using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using prjData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static OpenXmlDLLDotnetFramework.DLLTemplate;

namespace OpenXmlDll
{
    public class clsMisc
    {
        static uint uniqueId;
        public static void CopySlide(string sourceFile, string destinationFile, int insertPosition)
        {

            using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, false))
            using (PresentationDocument targetDoc = PresentationDocument.Open(destinationFile, true))
            {

                SlideId sourceSlideId = sourceDoc.PresentationPart.Presentation.SlideIdList.ChildElements.OfType<SlideId>().ElementAt(0);
                string sourceSlideRelId = sourceSlideId.RelationshipId;
                SlidePart sourceSlidePart = (SlidePart)sourceDoc.PresentationPart.GetPartById(sourceSlideRelId);
                SlidePart newSlidePart = targetDoc.PresentationPart.AddPart(sourceSlidePart);

                SlideIdList slideIdList = targetDoc.PresentationPart.Presentation.SlideIdList;
                SlideId newSlideId = new SlideId();
                newSlideId.RelationshipId = targetDoc.PresentationPart.GetIdOfPart(newSlidePart);

                uint maxSlideId = 1;
                foreach (SlideId slideId in slideIdList.ChildElements.OfType<SlideId>())
                {
                    if (slideId.Id > maxSlideId)
                    {
                        maxSlideId = slideId.Id;
                    }
                }
                newSlideId.Id = maxSlideId + 1;

                slideIdList.InsertAt(newSlideId, insertPosition);

                targetDoc.PresentationPart.Presentation.Save();
            }


        }


        public static void CopySlide2(string sourceFilePath, string destinationFilePath, int position)
        {
            using (PresentationDocument sourcePresentation = PresentationDocument.Open(sourceFilePath, false))
            using (PresentationDocument destinationPresentation = PresentationDocument.Open(destinationFilePath, true))
            { PresentationPart sourcePresentationPart = sourcePresentation.PresentationPart; PresentationPart destinationPresentationPart = destinationPresentation.PresentationPart; SlideId sourceSlideId = sourcePresentationPart.Presentation.SlideIdList.ChildElements.First() as SlideId; SlidePart sourceSlidePart = sourcePresentationPart.GetPartById(sourceSlideId.RelationshipId) as SlidePart; SlidePart newSlidePart = destinationPresentationPart.AddNewPart<SlidePart>(); sourceSlidePart.Slide.Save(newSlidePart); string slideMasterId = destinationPresentationPart.GetIdOfPart(newSlidePart); SlideId newSlideId = new SlideId { Id = destinationPresentationPart.Presentation.SlideIdList.ChildElements.Count > 0 ? (destinationPresentationPart.Presentation.SlideIdList.ChildElements.Cast<SlideId>().Max(s => s.Id.Value) + 1) : 256U, RelationshipId = slideMasterId }; SlideIdList slideIdList = destinationPresentationPart.Presentation.SlideIdList ?? new SlideIdList(); slideIdList.InsertAt(newSlideId, position); destinationPresentationPart.Presentation.SlideIdList = slideIdList; destinationPresentationPart.Presentation.Save(); }
        }


        public void InsertSlide(string sourceFilePath, string destinationFilePath, int position)
        {
            // Open the source and destination presentations
            using (PresentationDocument sourcePresentation = PresentationDocument.Open(sourceFilePath, false))
            using (PresentationDocument destinationPresentation = PresentationDocument.Open(destinationFilePath, true))
            {
                PresentationPart sourcePresentationPart = sourcePresentation.PresentationPart;
                PresentationPart destinationPresentationPart = destinationPresentation.PresentationPart;

                // Get the first slide from the source presentation
                SlideId sourceSlideId = sourcePresentationPart.Presentation.SlideIdList.ChildElements.First() as SlideId;
                SlidePart sourceSlidePart = sourcePresentationPart.GetPartById(sourceSlideId.RelationshipId) as SlidePart;

                // Add a new slide part to the destination presentation
                SlidePart newSlidePart = destinationPresentationPart.AddNewPart<SlidePart>();
                sourceSlidePart.Slide.Save(newSlidePart);

                // Create a new SlideId and add it to the SlideIdList
                SlideId newSlideId = new SlideId
                {
                    Id = destinationPresentationPart.Presentation.SlideIdList.ChildElements.Count > 0
                            ? (destinationPresentationPart.Presentation.SlideIdList.ChildElements.Cast<SlideId>().Max(s => s.Id.Value) + 1)
                            : 256U,
                    RelationshipId = destinationPresentationPart.GetIdOfPart(newSlidePart)
                };

                // Insert the new SlideId at the specified position
                destinationPresentationPart.Presentation.SlideIdList.InsertAt(newSlideId, position);

                // Save the destination presentation
                destinationPresentationPart.Presentation.Save();
            }
        }

        public void CopySlide(string sourceFile, string targetFile, int sourceSlideIndex, int targetPosition)
        {
            using (PresentationDocument sourcePresentation = PresentationDocument.Open(sourceFile, false))
            using (PresentationDocument targetPresentation = PresentationDocument.Open(targetFile, true))
            {
                PresentationPart sourcePresentationPart = sourcePresentation.PresentationPart;
                PresentationPart targetPresentationPart = targetPresentation.PresentationPart;

                // Get the source slide
                SlidePart sourceSlide = sourcePresentationPart.SlideParts.ElementAt(sourceSlideIndex);

                // Clone the slide and its related parts
                SlidePart newSlide = targetPresentationPart.AddPart<SlidePart>(sourceSlide);

                // Create a unique relationship ID for the new slide
                SlideIdList slideIdList = targetPresentationPart.Presentation.SlideIdList;
                uint maxSlideId = 6575;
                if (slideIdList != null && slideIdList.ChildElements.Count > 0)
                {
                    maxSlideId = slideIdList.ChildElements.Cast<SlideId>().Max(x => x.Id.Value) + 1;
                }

                // Create a new slide ID
                SlideId newSlideId = new SlideId();
                newSlideId.Id = maxSlideId;
                newSlideId.RelationshipId = targetPresentationPart.GetIdOfPart(newSlide);

                // Insert the new slide at the specified position
                slideIdList.InsertAt(newSlideId, targetPosition);

                // Save changes
                targetPresentationPart.Presentation.Save();
            }


        }


        public static void Copy(Stream sourcePresentationStream, uint copiedSlidePosition, Stream destPresentationStream)
        {
            using (var destDoc = PresentationDocument.Open(destPresentationStream, true))
            {
                var sourceDoc = PresentationDocument.Open(sourcePresentationStream, false);
                var destPresentationPart = destDoc.PresentationPart;
                var destPresentation = destPresentationPart.Presentation;

                var sourcePresentationPart = sourceDoc.PresentationPart;
                var sourcePresentation = sourcePresentationPart.Presentation;

                SlideIdList slideIdList = destPresentationPart.Presentation.SlideIdList;

                uint maxId = 6575;

                if (slideIdList != null && slideIdList.ChildElements.Count > 0)
                {
                    maxId = slideIdList.ChildElements.Cast<SlideId>().Max(x => x.Id.Value) + 1;
                }

                int copiedSlideIndex = (int)--copiedSlidePosition;

                int countSlidesInSourcePresentation = sourcePresentation.SlideIdList.Count();
                if (copiedSlideIndex < 0 || copiedSlideIndex >= countSlidesInSourcePresentation)
                    throw new ArgumentOutOfRangeException(nameof(copiedSlidePosition));

                SlideId copiedSlideId = sourcePresentationPart.Presentation.SlideIdList.ChildElements[copiedSlideIndex] as SlideId;
                SlidePart copiedSlidePart = sourcePresentationPart.GetPartById(copiedSlideId.RelationshipId) as SlidePart;

                SlidePart addedSlidePart = destPresentationPart.AddPart<SlidePart>(copiedSlidePart);

                SlideMasterPart addedSlideMasterPart = destPresentationPart.AddPart(addedSlidePart.SlideLayoutPart.SlideMasterPart);


                // Create new slide ID
                maxId++;
                SlideId slideId = new SlideId
                {
                    Id = maxId,
                    RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlidePart)
                };

                //_uniqueId = GetMaxIdFromChild(destPresentation.SlideMasterIdList);

                //destPresentation.SlideIdList.Append(slideId);

                //// Create new master slide ID
                //_uniqueId++;
                //SlideMasterId slideMaterId = new SlideMasterId
                //{
                //    Id = _uniqueId,
                //    RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlideMasterPart)
                //};
                //destDoc.PresentationPart.Presentation.SlideMasterIdList.Append(slideMaterId);

                // change slide layout ID
                //  FixSlideLayoutIds(destDoc.PresentationPart);

                destDoc.PresentationPart.Presentation.Save();
            }
            sourcePresentationStream.Close();
            destPresentationStream.Close();
        }


        public static void InsertSlideFromSourcePptToDestinationPpt(string sourcePptFile, string destinationPptFile, int slideIndexToInsert)
        {
            using (PresentationDocument sourcePptDoc = PresentationDocument.Open(sourcePptFile, true))
            {
                // Get the slide parts
                var slideParts = sourcePptDoc.PresentationPart.SlideParts;

                // Check if the slide index is valid
                if (slideIndexToInsert >= slideParts.Count())
                {
                    throw new IndexOutOfRangeException("Slide index is out of range");
                }

                // Get the slide part to insert
                SlidePart sourceSlidePart = slideParts.ElementAt(slideIndexToInsert);

                // Open destination PPT file
                using (PresentationDocument destinationPptDoc = PresentationDocument.Open(destinationPptFile, true))
                {
                    // Get the presentation part
                    PresentationPart presentationPart = destinationPptDoc.PresentationPart;

                    // Create a new slide part
                    SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();

                    // Copy the slide content from source to destination
                    using (Stream sourceStream = sourceSlidePart.GetStream())
                    using (Stream destinationStream = newSlidePart.GetStream())
                    {
                        sourceStream.CopyTo(destinationStream);
                    }

                    // Save changes to the destination PPT file
                    destinationPptDoc.Save();
                }
            }




        }


        public static void InsertSlide(string sourcePPT, string destinationPPT)

        {

            int id = 0;

            using (PresentationDocument destinationDoc = PresentationDocument.Open(destinationPPT, true))

            {

                PresentationPart destinationPPTPart = destinationDoc?.PresentationPart;

                if (destinationPPTPart.Presentation.SlideIdList == null)

                    destinationPPTPart.Presentation.SlideIdList = new SlideIdList();

                using (PresentationDocument sourceDoc = PresentationDocument.Open(sourcePPT, true))
                {
                    PresentationPart sourcePPTPart = sourceDoc.PresentationPart;

                    uniqueId = GetMaxSlideMasterId(destinationPPTPart.Presentation.SlideMasterIdList);

                    uint maxSlideId = GetMaxSlideId(destinationPPTPart.Presentation.SlideIdList);

                    foreach (SlideId slideId in sourcePPTPart.Presentation.SlideIdList)

                    {

                        id++;

                        SlidePart sourceSlidePart = (SlidePart)sourcePPTPart.GetPartById(slideId.RelationshipId);

                        string relId = "rel" + id;

                        SlidePart destinationSlidePart = destinationPPTPart?.AddPart<SlidePart>(sourceSlidePart, relId);

                        SlideMasterPart destinationMasterPart = destinationSlidePart?.SlideLayoutPart?.SlideMasterPart;

                        destinationPPTPart.AddPart(destinationMasterPart);

                        uniqueId++;

                        SlideMasterId newSlideMasterId = new SlideMasterId

                        {

                            RelationshipId = destinationPPTPart?.GetIdOfPart(destinationMasterPart),

                            Id = uniqueId

                        };

                        if (!destinationPPTPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>().Any(x => x.RelationshipId == newSlideMasterId.RelationshipId))

                        {

                            destinationPPTPart?.Presentation?.SlideMasterIdList.Append(newSlideMasterId);

                        }

                        maxSlideId++;

                        SlideId newSlideId = new SlideId

                        {

                            RelationshipId = relId,

                            Id = maxSlideId

                        };

                        InsertSlideAtIndex(destinationPPTPart.Presentation.SlideIdList, newSlideId);

                    }

                    FixSlideLayoutIds(destinationPPTPart);

                }

                destinationPPTPart.Presentation.Save();

            }

        }



        public static void InsertSlideAtIndex(SlideIdList slideIdList, SlideId newSlideId)

        {

            var slideIds = slideIdList?.Elements<SlideId>().ToList();

            slideIdList.Append(newSlideId);

        }

        static void FixSlideLayoutIds(PresentationPart presPart)

        {

            foreach (SlideMasterPart slideMasterPart in presPart?.SlideMasterParts)

            {

                foreach (SlideLayoutId slideLayoutId in slideMasterPart?.SlideMaster?.SlideLayoutIdList)

                {

                    uniqueId++;

                    slideLayoutId.Id = uniqueId;

                }

                slideMasterPart.SlideMaster.Save();

            }

        }

        public static uint GetMaxSlideMasterId(SlideMasterIdList slideMasterIdList)

        {

            uint max = 2147483648;

            if (slideMasterIdList != null)

            {

                foreach (SlideMasterId slideMasterId in slideMasterIdList?.Elements<SlideMasterId>())

                {

                    uint id = slideMasterId.Id;

                    if (id > max) max = id;

                }

            }

            return max;

        }

        public static uint GetMaxSlideId(SlideIdList slideIdList)
        {

            uint max = 256;

            if (slideIdList != null)

            {

                foreach (SlideId slide in slideIdList?.Elements<SlideId>())

                {

                    uint id = slide.Id;

                    if (id > max) max = id;

                }

            }

            return max;

        }





        //


        public void insertSlide()
        {



        }

        // string sourceFile = "C:\\Testing\\OpenXmlSlideSolutions\\CopySlidesFromBeginning\\CopySlidesFromListOfInput\\NewFolder\\fit_to_concept_template.pptx";
        //  string destinationFile = "C:\\Testing\\OpenXmlSlideSolutions\\CopySlidesFromBeginning\\CopySlidesFromListOfInput\\NewFolder\\OverallImpressionsTemplate.pptx";

        //   int[] sourceInput = { 1, 2, 3 }; MergeSlideWithSlideArray(sourceFile, destinationFile, sourceInput);


        public static void MergeSlideWithSlideArrayBrandexTesting(string sourcePresentation, string destPresentation, int[] insertIndex, int rId)
        {

            Random random = new Random();
            int randomNumber = random.Next(0, 1000);
            int id = rId + randomNumber;

            using (PresentationDocument destinationDoc = PresentationDocument.Open(destPresentation, true))
            {
                PresentationPart destinationPresPart = destinationDoc.PresentationPart;

                if (destinationPresPart.Presentation.SlideIdList == null)
                    destinationPresPart.Presentation.SlideIdList = new SlideIdList();

                using (PresentationDocument sourceDoc = PresentationDocument.Open(sourcePresentation, true))
                {
                    PresentationPart sourcePresPart = sourceDoc.PresentationPart;

                    uniqueId = GetMaxSlideMasterId(destinationPresPart.Presentation.SlideMasterIdList);
                    uint maxSlideId = GetMaxSlideId(destinationPresPart.Presentation.SlideIdList);

                    var sourceSlideIds = sourcePresPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

                    if (insertIndex.Length > sourceSlideIds.Count)
                        throw new ArgumentException("Not enough slides in source to match insert positions");

                    for (int i = 0; i < insertIndex.Length; i++)
                    {
                        id++;
                        SlideId sourceSlideId = sourceSlideIds[i];
                        SlidePart sourceSlidePart = (SlidePart)sourcePresPart.GetPartById(sourceSlideId.RelationshipId);

                        string relId = "rel" + id;

                        SlidePart destinationSlidePart = destinationPresPart.AddPart<SlidePart>(sourceSlidePart, relId);

                        SlideMasterPart destinationMasterPart = destinationSlidePart.SlideLayoutPart.SlideMasterPart;
                        destinationPresPart.AddPart(destinationMasterPart);

                        uniqueId++;
                        SlideMasterId newSlideMasterId = new SlideMasterId
                        {
                            RelationshipId = destinationPresPart.GetIdOfPart(destinationMasterPart),
                            Id = uniqueId
                        };

                        if (!destinationPresPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>().Any(x => x.RelationshipId == newSlideMasterId.RelationshipId))
                        {
                            destinationPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                        }

                        maxSlideId++;

                        SlideId newSlideId = new SlideId
                        {
                            RelationshipId = relId,
                            Id = maxSlideId
                        };

                        InsertSlideAtIndexArray(destinationPresPart.Presentation.SlideIdList, newSlideId, insertIndex[i] - 1);
                    }
                    FixSlideLayoutIds(destinationPresPart);
                }


                ChangePresentationPart(destinationPresPart);

                destinationPresPart.Presentation.Save();

                destinationDoc?.Dispose();



            }
        }


        public static void ChangePresentationPart(PresentationPart presentationPart1)
        {
            Presentation presentation1 = presentationPart1.Presentation;


            SlideMasterIdList slideMasterIdList1 = presentation1.GetFirstChild<SlideMasterIdList>();
            NotesMasterIdList notesMasterIdList1 = presentation1.GetFirstChild<NotesMasterIdList>();
            HandoutMasterIdList handoutMasterIdList1 = presentation1.GetFirstChild<HandoutMasterIdList>();
            SlideIdList slideIdList1 = presentation1.GetFirstChild<SlideIdList>();

            SlideMasterId slideMasterId1 = slideMasterIdList1.Elements<SlideMasterId>().ElementAt(1);
            slideMasterId1.RelationshipId = "Ra8b171aeae684ae7";

            SlideMasterId slideMasterId2 = new SlideMasterId() { Id = (UInt32Value)2147484007U, RelationshipId = "R28d28e75f45e47f9" };
            slideMasterIdList1.Append(slideMasterId2);

            SlideMasterId slideMasterId3 = new SlideMasterId() { Id = (UInt32Value)2147484008U, RelationshipId = "Rd619b3d348994ec2" };
            slideMasterIdList1.Append(slideMasterId3);

            SlideMasterId slideMasterId4 = new SlideMasterId() { Id = (UInt32Value)2147484009U, RelationshipId = "Rd8b4dcb04e6540b3" };
            slideMasterIdList1.Append(slideMasterId4);

            NotesMasterId notesMasterId1 = notesMasterIdList1.GetFirstChild<NotesMasterId>();
            notesMasterId1.Id = "rId103";

            HandoutMasterId handoutMasterId1 = handoutMasterIdList1.GetFirstChild<HandoutMasterId>();
            handoutMasterId1.Id = "rId104";

            SlideId slideId1 = slideIdList1.GetFirstChild<SlideId>();
            SlideId slideId2 = slideIdList1.Elements<SlideId>().ElementAt(1);
            SlideId slideId3 = slideIdList1.Elements<SlideId>().ElementAt(2);
            SlideId slideId4 = slideIdList1.Elements<SlideId>().ElementAt(3);
            SlideId slideId5 = slideIdList1.Elements<SlideId>().ElementAt(4);
            SlideId slideId6 = slideIdList1.Elements<SlideId>().ElementAt(5);
            SlideId slideId7 = slideIdList1.Elements<SlideId>().ElementAt(6);
            SlideId slideId8 = slideIdList1.Elements<SlideId>().ElementAt(7);
            SlideId slideId9 = slideIdList1.Elements<SlideId>().ElementAt(8);
            SlideId slideId10 = slideIdList1.Elements<SlideId>().ElementAt(9);
            SlideId slideId11 = slideIdList1.Elements<SlideId>().ElementAt(10);
            SlideId slideId12 = slideIdList1.Elements<SlideId>().ElementAt(11);
            SlideId slideId13 = slideIdList1.Elements<SlideId>().ElementAt(12);
            SlideId slideId14 = slideIdList1.Elements<SlideId>().ElementAt(13);
            SlideId slideId15 = slideIdList1.Elements<SlideId>().ElementAt(14);
            SlideId slideId16 = slideIdList1.Elements<SlideId>().ElementAt(15);
            SlideId slideId17 = slideIdList1.Elements<SlideId>().ElementAt(16);
            SlideId slideId18 = slideIdList1.Elements<SlideId>().ElementAt(17);
            SlideId slideId19 = slideIdList1.Elements<SlideId>().ElementAt(18);
            SlideId slideId20 = slideIdList1.Elements<SlideId>().ElementAt(19);
            SlideId slideId21 = slideIdList1.Elements<SlideId>().ElementAt(20);
            SlideId slideId22 = slideIdList1.Elements<SlideId>().ElementAt(21);
            SlideId slideId23 = slideIdList1.Elements<SlideId>().ElementAt(22);
            SlideId slideId24 = slideIdList1.Elements<SlideId>().ElementAt(23);
            SlideId slideId25 = slideIdList1.Elements<SlideId>().ElementAt(24);
            SlideId slideId26 = slideIdList1.Elements<SlideId>().ElementAt(25);
            SlideId slideId27 = slideIdList1.Elements<SlideId>().ElementAt(26);
            SlideId slideId28 = slideIdList1.Elements<SlideId>().ElementAt(27);
            SlideId slideId29 = slideIdList1.Elements<SlideId>().ElementAt(28);
            SlideId slideId30 = slideIdList1.Elements<SlideId>().ElementAt(29);
            SlideId slideId31 = slideIdList1.Elements<SlideId>().ElementAt(30);
            SlideId slideId32 = slideIdList1.Elements<SlideId>().ElementAt(31);
            SlideId slideId33 = slideIdList1.Elements<SlideId>().ElementAt(32);
            SlideId slideId34 = slideIdList1.Elements<SlideId>().ElementAt(33);
            SlideId slideId35 = slideIdList1.Elements<SlideId>().ElementAt(34);
            SlideId slideId36 = slideIdList1.Elements<SlideId>().ElementAt(35);
            SlideId slideId37 = slideIdList1.Elements<SlideId>().ElementAt(36);
            SlideId slideId38 = slideIdList1.Elements<SlideId>().ElementAt(37);
            SlideId slideId39 = slideIdList1.Elements<SlideId>().ElementAt(38);
            SlideId slideId40 = slideIdList1.Elements<SlideId>().ElementAt(39);
            SlideId slideId41 = slideIdList1.Elements<SlideId>().ElementAt(40);
            SlideId slideId42 = slideIdList1.Elements<SlideId>().ElementAt(41);
            SlideId slideId43 = slideIdList1.Elements<SlideId>().ElementAt(42);
            SlideId slideId44 = slideIdList1.Elements<SlideId>().ElementAt(43);
            SlideId slideId45 = slideIdList1.Elements<SlideId>().ElementAt(44);
            SlideId slideId46 = slideIdList1.Elements<SlideId>().ElementAt(45);
            SlideId slideId47 = slideIdList1.Elements<SlideId>().ElementAt(46);
            SlideId slideId48 = slideIdList1.Elements<SlideId>().ElementAt(47);
            SlideId slideId49 = slideIdList1.Elements<SlideId>().ElementAt(48);
            SlideId slideId50 = slideIdList1.Elements<SlideId>().ElementAt(49);
            SlideId slideId51 = slideIdList1.Elements<SlideId>().ElementAt(50);
            SlideId slideId52 = slideIdList1.Elements<SlideId>().ElementAt(51);
            SlideId slideId53 = slideIdList1.Elements<SlideId>().ElementAt(52);
            SlideId slideId54 = slideIdList1.Elements<SlideId>().ElementAt(53);
            SlideId slideId55 = slideIdList1.Elements<SlideId>().ElementAt(54);
            SlideId slideId56 = slideIdList1.Elements<SlideId>().ElementAt(55);
            SlideId slideId57 = slideIdList1.Elements<SlideId>().ElementAt(56);
            SlideId slideId58 = slideIdList1.Elements<SlideId>().ElementAt(57);
            SlideId slideId59 = slideIdList1.Elements<SlideId>().ElementAt(58);
            SlideId slideId60 = slideIdList1.Elements<SlideId>().ElementAt(59);
            SlideId slideId61 = slideIdList1.Elements<SlideId>().ElementAt(60);
            SlideId slideId62 = slideIdList1.Elements<SlideId>().ElementAt(61);
            SlideId slideId63 = slideIdList1.Elements<SlideId>().ElementAt(62);
            SlideId slideId64 = slideIdList1.Elements<SlideId>().ElementAt(63);
            SlideId slideId65 = slideIdList1.Elements<SlideId>().ElementAt(64);
            SlideId slideId66 = slideIdList1.Elements<SlideId>().ElementAt(65);
            SlideId slideId67 = slideIdList1.Elements<SlideId>().ElementAt(66);
            SlideId slideId68 = slideIdList1.Elements<SlideId>().ElementAt(67);
            SlideId slideId69 = slideIdList1.Elements<SlideId>().ElementAt(68);
            SlideId slideId70 = slideIdList1.Elements<SlideId>().ElementAt(69);
            SlideId slideId71 = slideIdList1.Elements<SlideId>().ElementAt(70);
            SlideId slideId72 = slideIdList1.Elements<SlideId>().ElementAt(71);
            SlideId slideId73 = slideIdList1.Elements<SlideId>().ElementAt(72);
            slideId1.Id = (UInt32Value)14323U;
            slideId1.RelationshipId = "rel852";
            slideId2.Id = (UInt32Value)14324U;
            slideId2.RelationshipId = "rel853";
            slideId3.Id = (UInt32Value)14325U;
            slideId3.RelationshipId = "rel854";
            slideId4.RelationshipId = "rId4";

            SlideId slideId74 = new SlideId() { Id = (UInt32Value)13896U, RelationshipId = "rId5" };
            slideIdList1.InsertBefore(slideId74, slideId5);

            SlideId slideId75 = new SlideId() { Id = (UInt32Value)13895U, RelationshipId = "rId6" };
            slideIdList1.InsertBefore(slideId75, slideId5);
            slideId5.Id = (UInt32Value)13898U;
            slideId5.RelationshipId = "rId7";
            slideId6.Id = (UInt32Value)13974U;
            slideId6.RelationshipId = "rId8";
            slideId7.Id = (UInt32Value)13901U;
            slideId7.RelationshipId = "rId9";
            slideId8.Id = (UInt32Value)13902U;
            slideId8.RelationshipId = "rId10";
            slideId9.Id = (UInt32Value)13941U;
            slideId9.RelationshipId = "rId11";
            slideId10.Id = (UInt32Value)13943U;
            slideId10.RelationshipId = "rId12";
            slideId11.Id = (UInt32Value)13942U;
            slideId11.RelationshipId = "rId13";
            slideId12.Id = (UInt32Value)13904U;
            slideId12.RelationshipId = "rId14";

            slideId13.Remove();
            slideId14.Id = (UInt32Value)13907U;
            slideId14.RelationshipId = "rId15";
            slideId15.Id = (UInt32Value)13905U;
            slideId15.RelationshipId = "rId16";
            slideId16.Id = (UInt32Value)911U;
            slideId16.RelationshipId = "rId17";
            slideId17.Id = (UInt32Value)13940U;
            slideId17.RelationshipId = "rId18";
            slideId18.Remove();
            slideId19.Remove();
            slideId20.Remove();
            slideId21.Remove();
            slideId22.Remove();
            slideId23.Id = (UInt32Value)13983U;
            slideId23.RelationshipId = "rId19";
            slideId24.Id = (UInt32Value)13977U;
            slideId24.RelationshipId = "rId20";

            SlideId slideId76 = new SlideId() { Id = (UInt32Value)13978U, RelationshipId = "rId21" };
            slideIdList1.InsertBefore(slideId76, slideId25);

            SlideId slideId77 = new SlideId() { Id = (UInt32Value)13986U, RelationshipId = "rId22" };
            slideIdList1.InsertBefore(slideId77, slideId25);

            SlideId slideId78 = new SlideId() { Id = (UInt32Value)13975U, RelationshipId = "rId23" };
            slideIdList1.InsertBefore(slideId78, slideId25);

            SlideId slideId79 = new SlideId() { Id = (UInt32Value)13906U, RelationshipId = "rId24" };
            slideIdList1.InsertBefore(slideId79, slideId25);

            SlideId slideId80 = new SlideId() { Id = (UInt32Value)13910U, RelationshipId = "rId25" };
            slideIdList1.InsertBefore(slideId80, slideId25);

            SlideId slideId81 = new SlideId() { Id = (UInt32Value)13911U, RelationshipId = "rId26" };
            slideIdList1.InsertBefore(slideId81, slideId25);
            slideId25.Id = (UInt32Value)13912U;
            slideId25.RelationshipId = "rId27";
            slideId26.Id = (UInt32Value)13913U;
            slideId26.RelationshipId = "rId28";
            slideId27.Id = (UInt32Value)13969U;
            slideId27.RelationshipId = "rId29";
            slideId28.Id = (UInt32Value)13914U;
            slideId28.RelationshipId = "rId30";
            slideId29.Id = (UInt32Value)13915U;
            slideId29.RelationshipId = "rId31";
            slideId30.Id = (UInt32Value)997U;
            slideId30.RelationshipId = "rId32";
            slideId31.Id = (UInt32Value)999U;
            slideId31.RelationshipId = "rId33";
            slideId32.Id = (UInt32Value)924U;
            slideId32.RelationshipId = "rId34";
            slideId33.Id = (UInt32Value)13995U;
            slideId33.RelationshipId = "rId35";
            slideId34.Id = (UInt32Value)13996U;
            slideId34.RelationshipId = "rId36";
            slideId35.Id = (UInt32Value)13922U;
            slideId35.RelationshipId = "rId37";
            slideId36.Id = (UInt32Value)13997U;
            slideId36.RelationshipId = "rId38";
            slideId37.Id = (UInt32Value)13998U;
            slideId37.RelationshipId = "rId39";
            slideId38.Id = (UInt32Value)13917U;
            slideId38.RelationshipId = "rId40";
            slideId39.Id = (UInt32Value)13999U;
            slideId39.RelationshipId = "rId41";
            slideId40.Id = (UInt32Value)13920U;
            slideId40.RelationshipId = "rId42";
            slideId41.Id = (UInt32Value)13927U;
            slideId41.RelationshipId = "rId43";
            slideId42.Id = (UInt32Value)14014U;
            slideId42.RelationshipId = "rId44";
            slideId43.Id = (UInt32Value)14015U;
            slideId43.RelationshipId = "rId45";
            slideId44.Id = (UInt32Value)14016U;
            slideId44.RelationshipId = "rId46";
            slideId45.Id = (UInt32Value)14322U;
            slideId45.RelationshipId = "rel104";
            slideId46.Id = (UInt32Value)13929U;
            slideId46.RelationshipId = "rId47";
            slideId47.Id = (UInt32Value)13931U;
            slideId47.RelationshipId = "rId49";
            slideId48.Id = (UInt32Value)13932U;
            slideId48.RelationshipId = "rId50";
            slideId49.Id = (UInt32Value)13971U;
            slideId49.RelationshipId = "rId51";

            SlideId slideId82 = new SlideId() { Id = (UInt32Value)13945U, RelationshipId = "rId52" };
            slideIdList1.InsertBefore(slideId82, slideId50);

            SlideId slideId83 = new SlideId() { Id = (UInt32Value)13980U, RelationshipId = "rId53" };
            slideIdList1.InsertBefore(slideId83, slideId50);

            SlideId slideId84 = new SlideId() { Id = (UInt32Value)14312U, RelationshipId = "rId54" };
            slideIdList1.InsertBefore(slideId84, slideId50);

            SlideId slideId85 = new SlideId() { Id = (UInt32Value)14313U, RelationshipId = "rId55" };
            slideIdList1.InsertBefore(slideId85, slideId50);

            SlideId slideId86 = new SlideId() { Id = (UInt32Value)14321U, RelationshipId = "rId56" };
            slideIdList1.InsertBefore(slideId86, slideId50);

            SlideId slideId87 = new SlideId() { Id = (UInt32Value)14314U, RelationshipId = "rId57" };
            slideIdList1.InsertBefore(slideId87, slideId50);

            SlideId slideId88 = new SlideId() { Id = (UInt32Value)14315U, RelationshipId = "rId58" };
            slideIdList1.InsertBefore(slideId88, slideId50);

            SlideId slideId89 = new SlideId() { Id = (UInt32Value)14316U, RelationshipId = "rId59" };
            slideIdList1.InsertBefore(slideId89, slideId50);

            SlideId slideId90 = new SlideId() { Id = (UInt32Value)14318U, RelationshipId = "rId60" };
            slideIdList1.InsertBefore(slideId90, slideId50);

            SlideId slideId91 = new SlideId() { Id = (UInt32Value)14319U, RelationshipId = "rId61" };
            slideIdList1.InsertBefore(slideId91, slideId50);

            SlideId slideId92 = new SlideId() { Id = (UInt32Value)14320U, RelationshipId = "rId62" };
            slideIdList1.InsertBefore(slideId92, slideId50);

            SlideId slideId93 = new SlideId() { Id = (UInt32Value)13985U, RelationshipId = "rId63" };
            slideIdList1.InsertBefore(slideId93, slideId50);

            SlideId slideId94 = new SlideId() { Id = (UInt32Value)13981U, RelationshipId = "rId64" };
            slideIdList1.InsertBefore(slideId94, slideId50);
            slideId50.Id = (UInt32Value)13921U;
            slideId50.RelationshipId = "rId65";
            slideId51.Id = (UInt32Value)13989U;
            slideId51.RelationshipId = "rId66";

            slideId52.Remove();
            slideId53.Remove();
            slideId54.Remove();
            slideId55.Remove();
            slideId56.Remove();
            slideId57.Remove();
            slideId58.Remove();
            slideId59.Remove();
            slideId60.Remove();
            slideId61.Remove();
            slideId62.Remove();
            slideId63.Remove();
            slideId64.Remove();
            slideId65.Id = (UInt32Value)13923U;
            slideId65.RelationshipId = "rId67";
            slideId66.Remove();
            slideId67.Remove();
            slideId68.Remove();
            slideId69.Remove();
            slideId70.Id = (UInt32Value)13924U;
            slideId70.RelationshipId = "rId68";
            slideId71.Id = (UInt32Value)13950U;
            slideId71.RelationshipId = "rId69";
            slideId72.Id = (UInt32Value)13918U;
            slideId72.RelationshipId = "rId70";
            slideId73.Id = (UInt32Value)14000U;
            slideId73.RelationshipId = "rId71";

            SlideId slideId95 = new SlideId() { Id = (UInt32Value)13925U, RelationshipId = "rId72" };
            slideIdList1.Append(slideId95);

            SlideId slideId96 = new SlideId() { Id = (UInt32Value)13948U, RelationshipId = "rId102" };
            slideIdList1.Append(slideId96);


            presentationPart1.Presentation.Save();
        }


        public static void MergeSlideWithSlideArray(string sourcePresentation, string destPresentation, int[] insertIndex, int rId)
        {
            Random random = new Random();
            int randomNumber = random.Next(0, 1000);
            int id = rId + randomNumber;

            try
            {
                using (PresentationDocument destinationDoc = PresentationDocument.Open(destPresentation, true))
                {
                    PresentationPart destinationPresPart = destinationDoc?.PresentationPart;

                    if (destinationPresPart.Presentation.SlideIdList == null)
                        destinationPresPart.Presentation.SlideIdList = new SlideIdList();

                    using (PresentationDocument sourceDoc = PresentationDocument.Open(sourcePresentation, true))
                    {
                        PresentationPart sourcePresPart = sourceDoc?.PresentationPart;

                        uniqueId = GetMaxSlideMasterId(destinationPresPart.Presentation.SlideMasterIdList);
                        uint maxSlideId = GetMaxSlideId(destinationPresPart.Presentation.SlideIdList);

                        var sourceSlideIds = sourcePresPart?.Presentation?.SlideIdList?.Elements<SlideId>().ToList();

                        if (insertIndex.Length > sourceSlideIds.Count)
                            throw new ArgumentException("Not enough slides in source to match insert positions");

                        for (int i = 0; i < insertIndex.Length; i++)
                        {
                            id++;
                            SlideId sourceSlideId = sourceSlideIds[i];
                            SlidePart sourceSlidePart = (SlidePart)sourcePresPart.GetPartById(sourceSlideId.RelationshipId);

                            string relId = "rel" + id;

                            SlidePart destinationSlidePart = destinationPresPart.AddPart<SlidePart>(sourceSlidePart, relId);

                            SlideMasterPart destinationMasterPart = destinationSlidePart?.SlideLayoutPart?.SlideMasterPart;
                            destinationPresPart.AddPart(destinationMasterPart);

                            uniqueId++;
                            SlideMasterId newSlideMasterId = new SlideMasterId
                            {
                                RelationshipId = destinationPresPart.GetIdOfPart(destinationMasterPart),
                                Id = uniqueId
                            };

                            if (!destinationPresPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>().Any(x => x.RelationshipId == newSlideMasterId.RelationshipId))
                            {
                                destinationPresPart?.Presentation?.SlideMasterIdList.Append(newSlideMasterId);
                            }

                            maxSlideId++;

                            SlideId newSlideId = new SlideId
                            {
                                RelationshipId = relId,
                                Id = maxSlideId
                            };

                            InsertSlideAtIndexArray(destinationPresPart.Presentation.SlideIdList, newSlideId, insertIndex[i] - 1);
                        }
                        FixSlideLayoutIds(destinationPresPart);
                    }
                    destinationPresPart.Presentation.Save();

                    destinationDoc?.Dispose();



                }
            }
            catch (Exception)
            {
                throw;
            }

            //Random random = new Random();
            //int randomNumber = random.Next(0, 1000);
            //int id = rId + randomNumber;

            //using (PresentationDocument destinationDoc = PresentationDocument.Open(destPresentation, true))
            //{
            //    PresentationPart destinationPresPart = destinationDoc.PresentationPart;

            //    if (destinationPresPart.Presentation.SlideIdList == null)
            //        destinationPresPart.Presentation.SlideIdList = new SlideIdList();

            //    using (PresentationDocument sourceDoc = PresentationDocument.Open(sourcePresentation, true))
            //    {
            //        PresentationPart sourcePresPart = sourceDoc.PresentationPart;

            //        uniqueId = GetMaxSlideMasterId(destinationPresPart.Presentation.SlideMasterIdList);
            //       uint maxSlideId = GetMaxSlideId(destinationPresPart.Presentation.SlideIdList);

            //        var sourceSlideIds = sourcePresPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

            //        if (insertIndex.Length > sourceSlideIds.Count)
            //            throw new ArgumentException("Not enough slides in source to match insert positions");

            //        for (int i = 0; i < insertIndex.Length; i++)
            //        {
            //            id++;
            //            SlideId sourceSlideId = sourceSlideIds[i];
            //            SlidePart sourceSlidePart = (SlidePart)sourcePresPart.GetPartById(sourceSlideId.RelationshipId);

            //            string relId = "rel" + id;

            //            SlidePart destinationSlidePart = destinationPresPart.AddPart<SlidePart>(sourceSlidePart, relId);

            //            SlideMasterPart destinationMasterPart = destinationSlidePart.SlideLayoutPart.SlideMasterPart;
            //            destinationPresPart.AddPart(destinationMasterPart);

            //            uniqueId++;
            //            SlideMasterId newSlideMasterId = new SlideMasterId
            //            {
            //                RelationshipId = destinationPresPart.GetIdOfPart(destinationMasterPart),
            //                Id = uniqueId
            //            };

            //            if (!destinationPresPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>().Any(x => x.RelationshipId == newSlideMasterId.RelationshipId))
            //            {
            //                destinationPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
            //            }

            //            maxSlideId++;

            //            SlideId newSlideId = new SlideId
            //            {
            //                RelationshipId = relId,
            //                Id = maxSlideId
            //            };

            //            InsertSlideAtIndexArray(destinationPresPart.Presentation.SlideIdList, newSlideId, insertIndex[i]-1);
            //        }
            //        FixSlideLayoutIds(destinationPresPart);
            //    }
            //    destinationPresPart.Presentation.Save();

            //    destinationDoc?.Dispose();



            //}
        }

        public static void InsertSlideAtIndexArray(SlideIdList slideIdList, SlideId newSlideId, int index)
        {
            try
            {
                var slideIds = slideIdList.Elements<SlideId>().ToList();

                if (index < 0 || index >= slideIds.Count)
                {
                    slideIdList.Append(newSlideId); // Add to the end if index is out of range
                }
                else
                {
                    var targetSlide = slideIds.ElementAt(index);

                    // Insert the new slide before the target slide
                    targetSlide.InsertBeforeSelf(newSlideId);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        static bool DoesSlideExist(PresentationDocument presentationDoc, int slideIndex)
        {
            try
            {
                // Get the SlideIdList from the PresentationPart
                var slideIdList = presentationDoc?.PresentationPart?.Presentation?.SlideIdList;

                // Check if the slide index is valid and the SlideId exists
                return slideIdList != null && slideIndex > 0 && slideIndex <= slideIdList.Count();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void DeleteSlideFromPPT(string filePath, int slideNo)
        {
            try
            {
                using (PresentationDocument destinationDoc = PresentationDocument.Open(filePath, true))
                {
                    if (DoesSlideExist(destinationDoc, slideNo))
                    {
                        PresentationPart presentationPart = destinationDoc?.PresentationPart;

                        // Get the presentation from the presentation part
                        Presentation presentation = presentationPart?.Presentation;

                        // Get the slide ID list
                        SlideIdList slideIdList = presentation?.SlideIdList;

                        // Get the slide ID of the specified slide
                        SlideId slideId = slideIdList?.ChildElements[slideNo] as SlideId;

                        // Get the relationship ID of the slide
                        string slideRelId = slideId?.RelationshipId;

                        // Remove the slide from the slide list
                        slideIdList.RemoveChild(slideId);

                        // Remove references to the slide from custom shows
                        if (presentation.CustomShowList != null)
                        {
                            foreach (CustomShow customShow in presentation.CustomShowList.Elements<CustomShow>())
                            {
                                if (customShow.SlideList != null)
                                {
                                    SlideListEntry entry = customShow.SlideList.ChildElements
                                        .FirstOrDefault(s => ((SlideListEntry)s).Id == slideRelId) as SlideListEntry;
                                    if (entry != null)
                                    {
                                        customShow.SlideList.RemoveChild(entry);
                                    }
                                }
                            }
                        }

                        // Save the modified presentation
                        presentation?.Save();

                        // Remove the slide part
                        destinationDoc?.PresentationPart?.DeletePart(slideRelId);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }



        public static async Task<string> DeleteSlideFromPPTAsync1(string filePath, int slideNo)
        {
            try
            {
                string res = "";


                while (!IsFileInUse(filePath))
                {
                    try
                    {
                        await Task.Run(() => DeleteSlideFromPPT(filePath, slideNo));
                        res = "Deletion Sucessful";
                        break;
                    }
                    catch (IOException)
                    {
                        // If the file is in use, wait and retry
                        //Console.WriteLine("File is in use. Waiting...");
                        Thread.Sleep(300); // Wait for 1 second before retrying
                    }
                }

                return res;
            }
            catch (Exception)
            {
                throw;
            }
            // bool fileOpened = IsFileInUse(filePath);
        }

        public static async Task<string> MergeSlideWithSlideArrayAsync1(string sourcePresentation, string destPresentation, int[] insertIndex, int rId)
        {

            try
            {
                //bool fileOpened = IsFileInUse(sourcePresentation) && IsFileInUse(destPresentation);

                string res = "";

                while (!IsFileInUse(destPresentation))
                {
                    try
                    {
                        await Task.Run(() => MergeSlideWithSlideArray(sourcePresentation, destPresentation, insertIndex, rId)); ;
                        res = "Merge Sucessful";
                        break;
                    }
                    catch (IOException)
                    {
                        // If the file is in use, wait and retry
                        //Console.WriteLine("File is in use. Waiting...");
                        Thread.Sleep(300); // Wait for 1 second before retrying
                    }
                }

                return res;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static bool IsFileInUse(string filePath)
        {
            try
            {
                // Try to open the file with exclusive access
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // If successful, the file is not in use
                    return false;
                }
            }
            catch (IOException)
            {
                // If an IOException is thrown, the file is in use
                return true;
            }
        }





        public static SlidePart GetSlidePartByIndex(PresentationPart presentationPart, int slideIndex)
        {
            try
            {
                // Get the list of slide IDs
                var slideIds = presentationPart.Presentation.SlideIdList.ChildElements;

                if (slideIndex >= 0 && slideIndex < slideIds.Count)
                {
                    // Get the slide ID at the specified index
                    var slideId = (SlideId)slideIds[slideIndex];

                    // Get the slide part by relationship ID
                    return (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                }

                return null; // Slide index out of range
            }
            catch (Exception)
            {
                throw;
            }


        }

        public static void ReplaceTextInSlide(Slide slide, string textToFind, string textToReplace)
        {
            try
            {
                var x = slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.CommonSlideData>();
                if (x != null)
                {
                    var y = x.GetFirstChild<DocumentFormat.OpenXml.Presentation.ShapeTree>();
                    if (y != null)
                    {
                        var shapes = y.Elements<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
                        if (shapes != null)
                        {
                            foreach (var s in shapes)
                            {
                                var text = s.GetFirstChild<DocumentFormat.OpenXml.Presentation.TextBody>();
                                if (text != null)
                                {
                                    var para = text.GetFirstChild<DocumentFormat.OpenXml.Drawing.Paragraph>();
                                    if (para != null)
                                    {
                                        var run = para.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();
                                        if (run != null)
                                        {
                                            foreach (var r in run)
                                            {
                                                var t = r.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                                                if (t != null)
                                                {
                                                    if (!(string.IsNullOrEmpty(t.Text) || string.IsNullOrWhiteSpace(t.Text)))
                                                    {
                                                        if (t.Text.Contains(textToFind))
                                                        {
                                                            t.Text = t.Text.Replace(textToFind, textToReplace);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //skip
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception)
            {
                throw;
            }



            //try
            //{

            //    // Iterate through all shapes in the slide
            //    foreach (var shape in slide?.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
            //    {
            //        // Check if the shape has text
            //        if (shape.TextBody != null)
            //        {
            //            // Iterate through all paragraphs in the shape
            //            foreach (var paragraph in shape?.TextBody?.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            //            {
            //                if (paragraph != null)
            //                {
            //                    // Iterate through all runs (text elements) in the paragraph
            //                    foreach (var run in paragraph?.Descendants<DocumentFormat.OpenXml.Drawing.Run>())
            //                    {
            //                        if (run != null)
            //                        {
            //                            // Check if the run contains the text to find
            //                            if (run.Text.Text.Contains(textToFind))
            //                            {
            //                                // Replace the text
            //                                run.Text.Text = run.Text.Text.Replace(textToFind, textToReplace);
            //                            }
            //                            else
            //                            {
            //                                //skip
            //                            }
            //                        }
            //                        else
            //                        {
            //                            //skip
            //                        }
            //                    }
            //                }
            //                else
            //                {
            //                    //skip
            //                }
            //            }
            //        }
            //        else
            //        {
            //            //skip
            //        }
            //    }
            //}
            //catch (Exception)
            //{

            //    throw;
            //}

        }



        public static void InsertPresentationAtPosition(string targetPptx, string sourcePptx, int position)
        {
            // Open the target presentation
            using (PresentationDocument targetDoc = PresentationDocument.Open(targetPptx, true))
            {
                // Open the source presentation
                using (PresentationDocument sourceDoc = PresentationDocument.Open(sourcePptx, false))
                {
                    // Get the slide parts from the source presentation
                    var sourceSlides = sourceDoc.PresentationPart.SlideParts.ToList();

                    // Get the presentation part of the target document
                    var targetPresentationPart = targetDoc.PresentationPart;

                    // Get the slide list from the target presentation
                    var targetSlides = targetPresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

                    // Make sure the position is valid
                    if (position < 0 || position > targetSlides.Count)
                    {
                        throw new ArgumentOutOfRangeException(nameof(position), "Invalid position to insert slides.");
                    }

                    // Insert the slides from the source presentation at the given position
                    foreach (var slidePart in sourceSlides)
                    {
                        // Clone the slide part to add it to the target presentation
                        var newSlidePart = targetPresentationPart.AddNewPart<SlidePart>();

                        // Copy the slide content from the source to the new slide part
                        newSlidePart.FeedData(slidePart.GetStream());

                        // Create a new slide ID for this slide
                        uint slideId = targetSlides.Count > 0 ? targetSlides.Last().Id.Value + 1 : 256;

                        // Add the slide to the target presentation's slideIdList
                        targetSlides.Insert(position, new SlideId() { Id = slideId, RelationshipId = targetPresentationPart.GetIdOfPart(newSlidePart) });

                        position++; // Increment position for each inserted slide
                    }

                    // Update the target presentation's slideIdList
                    targetPresentationPart.Presentation.SlideIdList.RemoveAllChildren();
                    foreach (var slideId in targetSlides)
                    {
                        targetPresentationPart.Presentation.SlideIdList.Append(slideId);
                    }

                    // Save the changes to the target presentation
                    targetPresentationPart.Presentation.Save();
                }
            }
        }




        public static void repTextInSlide(string filePath, string textToFind, string textToReplace, int targetSlideIndex)
        {

            //string filePath = @"C:\excelfiles\RACKEM\Final\MRRxNaming.pptx";
            //string textToFind = "<<AttributeEvaluationTitle2>>";

            // string textToReplace = "Attribute #2";
            // int targetSlideIndex = 38; // Slide index (0-based39) to modify


            try
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Open(filePath, true))
                {
                    // Get the presentation part
                    PresentationPart presentationPart = presentationDoc.PresentationPart;

                    // Get the slide at the specified index
                    SlidePart slidePart = GetSlidePartByIndex(presentationPart, targetSlideIndex);

                    if (slidePart != null)
                    {
                        // Get the slide
                        Slide slide = slidePart.Slide;

                        // Find and replace text in the slide
                        ReplaceTextInSlide(slide, textToFind, textToReplace);

                        // Save the changes
                        presentationDoc.Save();

                    }

                }
            }
            catch (Exception)
            {
                throw;
            }

        }




        public static string getAttributeTitle(string project, string chart)
        {
            string repText = "";
            System.Data.DataTable dt1 = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getAttributeEvaluationTitle] " + "'" + project + "'," + "'" + chart + "'");

            foreach (DataRow row in dt1.Rows)
            {
                repText = Convert.ToString(row["strAttributeEvaluationTitle"]);
            }

            return repText;


        }



        public static async Task repTextInSlideAsync(string filePath, string textToFind, string textToReplace, int targetSlideIndex)
        {
            try
            {

                bool fileOpened = false;

                //string res = "";

                while (!IsFileInUse(filePath))
                {
                    try
                    {
                        await Task.Run(() => repTextInSlide(filePath, textToFind, textToReplace, targetSlideIndex));

                        //res = "Merge Sucessful";
                        break;
                    }
                    catch (IOException)
                    {
                        // If the file is in use, wait and retry
                        //Console.WriteLine("File is in use. Waiting...");
                        Thread.Sleep(300); // Wait for 1 second before retrying
                    }
                }

            }
            catch (Exception)
            {
                throw;
            }




        }

        public static void replaceText(string filePath, string searchText, string replaceText)
        {


            // Open the PowerPoint presentation
            using (PresentationDocument presentationDoc = PresentationDocument.Open(filePath, true))
            {
                // Get the presentation part
                PresentationPart presentationPart = presentationDoc.PresentationPart;

                // Iterate through all slides
                foreach (SlidePart slidePart in presentationPart.SlideParts)
                {
                    // Replace text in the slide
                    ReplaceTextInSlide2(slidePart, searchText, replaceText);
                }

                // Save the changes
                presentationDoc.Save();
            }

        }


        static void ReplaceTextInSlide2(SlidePart slidePart, string searchText, string replaceText)
        {
            // Get the slide content
            Slide slide = slidePart.Slide;

            // Iterate through all paragraphs in the slide
            foreach (var paragraph in slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                // Iterate through all text elements in the paragraph
                foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                {
                    // Replace the text if it matches the search text
                    if (text.Text.Contains(searchText))
                    {
                        text.Text = text.Text.Replace(searchText, replaceText);
                    }
                }
            }
        }









        //public static void ReplaceTextInPresentation(string filePath, string searchText, string replaceText, int slideIndex)
        //{
        //    using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
        //    {
        //        PresentationPart presentationPart = presentationDocument.PresentationPart;
        //        SlidePart slidePart = presentationPart.SlideParts.ElementAt(slideIndex);

        //        foreach (A.Paragraph paragraph in slidePart.Slide.Descendants<A.Paragraph>())
        //        {
        //            foreach (A.Run run in paragraph.Elements<A.Run>())
        //            {
        //                if (run.InnerText.Contains(searchText))
        //                {
        //                    A.Text text = run.GetFirstChild<A.Text>();
        //                    if (text != null)
        //                    {
        //                        text.Text = text.Text.Replace(searchText, replaceText);
        //                    }
        //                }
        //            }
        //        }

        //        slidePart.Slide.Save();
        //    }
        //}







    }
}






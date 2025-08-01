using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace ExcelChartsBlazorOpenxml.Components
{
    public static class MergeDeleteHelperMethods
    {


        public static int id = 20;

        public static void MergeSlideWithSlideArray(string sourcePresentation, string destPresentation, int[] insertIndex)
        {
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

                        //SlidePart destinationSlidePart = destinationPresPart?.AddPart<SlidePart>(sourceSlidePart, relId);

                        SlidePart destinationSlidePart = destinationPresPart.AddPart<SlidePart>(sourceSlidePart, relId);
                        //foreach (ChartPart chartPart in sourceSlidePart.ChartParts)
                        //{
                        //    ChartPart newChartPart = destinationSlidePart.AddPart(chartPart);
                        //}

                        SlideMasterPart destinationMasterPart = destinationSlidePart.SlideLayoutPart.SlideMasterPart;
                        destinationPresPart.AddPart(destinationMasterPart);

                        uniqueId++;

                        var x = destinationPresPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>().ToList();

                        var y = destinationPresPart.SlideMasterParts;

                        var z = destinationPresPart.Presentation.SlideIdList;

                        SlideMasterId newSlideMasterId = new SlideMasterId
                        {
                            RelationshipId = destinationPresPart.GetIdOfPart(destinationMasterPart),
                            Id = uniqueId
                        };


                        if (!destinationPresPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>().Any(x => x.RelationshipId == newSlideMasterId.RelationshipId))
                        {
                            destinationPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                        }

                        //if (!destinationPresPart.SlideMasterParts.Any(m => m.Uri == sourceSlidePart.SlideLayoutPart?.SlideMasterPart?.Uri))
                        //{
                        //    SlideMasterPart newMasterPart = destinationPresPart.AddPart(sourceSlidePart.SlideLayoutPart.SlideMasterPart);
                        //}

                        maxSlideId++;

                        SlideId newSlideId = new SlideId
                        {
                            RelationshipId = relId,
                            Id = maxSlideId
                        };

                        InsertSlideAtIndexArray(destinationPresPart.Presentation.SlideIdList, newSlideId, insertIndex[i]);
                    }
                    FixSlideLayoutIds(destinationPresPart);
                }


                destinationPresPart.Presentation.Save();
            }
        }


        static uint uniqueId = 2647484033;

        static void FixSlideLayoutIds(PresentationPart presPart)
        {
            foreach (SlideMasterPart slideMasterPart in presPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    uniqueId++;
                    slideLayoutId.Id = (uint)uniqueId;
                }

                slideMasterPart.SlideMaster.Save();
            }
        }

        public static uint GetMaxSlideMasterId(SlideMasterIdList slideMasterIdList)
        {
            uint max = 2147483648;

            if (slideMasterIdList != null)
            {
                foreach (SlideMasterId child in slideMasterIdList.Elements<SlideMasterId>())
                {
                    uint id = child.Id;

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
                foreach (SlideId child in slideIdList.Elements<SlideId>())
                {
                    uint id = child.Id;

                    if (id > max)
                        max = id;
                }
            }
            return max;
        }

        public static void InsertSlideAtIndexArray(SlideIdList slideIdList, SlideId newSlideId, int index)
        {
            var slideIds = slideIdList.Elements<SlideId>().ToList();

            if (index < 0 || index >= slideIds.Count)
            {
                slideIdList.Append(newSlideId); // Add to the end if index is out of range
            }
            else
            {
                var targetSlide = slideIds.ElementAt(index);
                targetSlide.InsertBeforeSelf(newSlideId);
            }

            // // Sort slides based on their IDs to ensure correct order
            // var sortedSlides = slideIdList.Elements<SlideId>()
            //     .OrderBy(slide => slide.Id)
            //     .ToList();

            // slideIdList.RemoveAllChildren<SlideId>();
            // foreach (var slide in sortedSlides)
            // {
            //     slideIdList.Append(slide);
            // }
        }


        static int CountSlides(string presentationFile)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation to the next CountSlide method
                // and return the slide count.
                return CountSlides(presentationDocument);
            }
        }


        static int CountSlides(PresentationDocument presentationDocument)
        {
            if (presentationDocument is null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;

            // Get the presentation part of document.
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // Get the slide count from the SlideParts.
            if (presentationPart is not null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }

            // Return the slide count to the previous method.
            return slidesCount;
        }

        static void DeleteSlide(string presentationFile, int slideIndex)
        {
            // Open the source document as read/write.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                // Pass the source document and the index of the slide to be deleted to the next DeleteSlide method.
                DeleteSlide(presentationDocument, slideIndex);
            }
        }

        static void DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
        {
            if (presentationDocument is null)
            {
                throw new ArgumentNullException(nameof(presentationDocument));
            }

            // Use the CountSlides sample to get the number of slides in the presentation.
            int slidesCount = CountSlides(presentationDocument);

            if (slideIndex < 0 || slideIndex >= slidesCount)
            {
                throw new ArgumentOutOfRangeException("slideIndex");
            }

            // Get the presentation part from the presentation document.
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // Get the presentation from the presentation part.
            Presentation? presentation = presentationPart?.Presentation;

            // Get the list of slide IDs in the presentation.
            SlideIdList? slideIdList = presentation?.SlideIdList;

            // Get the slide ID of the specified slide
            SlideId? slideId = slideIdList?.ChildElements[slideIndex] as SlideId;

            // Get the relationship ID of the slide.
            string? slideRelId = slideId?.RelationshipId;

            // If there's no relationship ID, there's no slide to delete.
            if (slideRelId is null)
            {
                return;
            }

            // Remove the slide from the slide list.
            slideIdList!.RemoveChild(slideId);

            // Remove references to the slide from all custom shows.
            if (presentation!.CustomShowList is not null)
            {
                // Iterate through the list of custom shows.
                foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList is not null)
                    {
                        // Declare a link list of slide list entries.
                        LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                        foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                        {
                            // Find the slide reference to remove from the custom show.
                            if (slideListEntry.Id is not null && slideListEntry.Id == slideRelId)
                            {
                                slideListEntries.AddLast(slideListEntry);
                            }
                        }

                        // Remove all references to the slide from the custom show.
                        foreach (SlideListEntry slideListEntry in slideListEntries)
                        {
                            customShow.SlideList.RemoveChild(slideListEntry);
                        }
                    }
                }
            }

            // Get the slide part for the specified slide.
            SlidePart slidePart = (SlidePart)presentationPart!.GetPartById(slideRelId);

            // Remove the slide part.
            presentationPart.DeletePart(slidePart);
        }




    }
}

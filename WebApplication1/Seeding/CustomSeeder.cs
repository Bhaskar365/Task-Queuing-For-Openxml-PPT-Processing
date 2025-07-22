using SharedModels;
using WebApplication1.Context;

namespace WebApplicationAPI.Seeding
{
    public class CustomSeeder
    {
        private readonly AppDBContext _context;

        public CustomSeeder(AppDBContext context)
        {
            _context = context;
        }

        public async Task SeedAsync()
        {
            if (!_context.FitToConceptTestTable.Any())
            {
                var data = new List<FitToConceptModel>
                {
                    new FitToConceptModel {ProjectName="RACKEM",TestName="UPTERLO",TestNameColor="#808080",Average=5.21,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType="FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMCLYDIN",TestNameColor="#FF0000",Average=5.19,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType="FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMFYDANT",TestNameColor="#808080",Average=5.17,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType="FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="YESIBY",TestNameColor="#808080",Average=5.15,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType="FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMLIVLY",TestNameColor="#FF0000",Average=5.15,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMROBLY",TestNameColor="#808080",Average=5.12,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMJORNY",TestNameColor="#FF0000",Average=5.11,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="PAMZELBY",TestNameColor="#FF0000",Average=5.10,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="RYDELTYX",TestNameColor="#FF0000",Average=5.10,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMROBVI",TestNameColor="#FF0000",Average=5.09,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="MIVZELTY",TestNameColor="#FF0000",Average=5.09,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="CYMDAZO",TestNameColor="#FF0000",Average=5.07,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMPZELTIV",TestNameColor="#FF0000",Average=5.07,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="PAMTIVGO",TestNameColor="#808080",Average=5.07,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="PAMSOLRIS",TestNameColor="#FF0000",Average=5.06,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="PAMVYA",TestNameColor="#FF0000",Average=5.02,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="BLYTRAVE",TestNameColor="#808080",Average=5.01,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="REMPOJI",TestNameColor="#FF0000",Average=5.00,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMCYZRA",TestNameColor="#FF0000",Average=4.98,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="EMKELVIQ",TestNameColor="#009900",Average=4.97,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="HAVUMFOR",TestNameColor="#808080",Average=4.97,TestNameBold=true,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="HEVKESSI",TestNameColor="#FF0000",Average=4.94,TestNameBold=false,AverageColor="#1D3C7C", ProjectTemplateType = "FittoConcept"},
                    new FitToConceptModel {ProjectName="RACKEM",TestName="[Refreshed Hist. Mean]",TestNameColor="#000000",Average=4.17,TestNameBold=true,AverageColor="#C2D7F2" , ProjectTemplateType="FittoConcept"}
               };

                await _context.FitToConceptTestTable.AddRangeAsync(data);
                await _context.SaveChangesAsync();
            }

            if (!_context.OverallImpressionsTestTable.Any())
            {
                var data = new List<OverallImpressionsModel>
                {
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "UPTERLO", Positive = 28.34, Neutral = 70.59, Negative = 1.07, TestNameBold = true, TestNameColor = "#808080", ProjectTemplateType="OverallImpressions" },
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "YESIBY", Positive = 27.81, Neutral = 68.72, Negative = 3.48, TestNameBold = true, TestNameColor = "#808080" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "CYMDAZO", Positive = 26.20, Neutral = 72.19, Negative = 1.60, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMJORNY", Positive = 25.94, Neutral = 71.66, Negative = 2.41, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "PAMZELBY", Positive = 25.13, Neutral = 73.53, Negative = 1.34, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "PAMSOLRIS", Positive = 24.87, Neutral = 73.26, Negative = 1.87, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "PAMVYA", Positive = 24.60, Neutral = 72.73, Negative = 2.67, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMROBLY", Positive = 24.06, Neutral = 74.06, Negative = 1.87, TestNameBold = true, TestNameColor = "#808080" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMFYDANT", Positive = 23.80, Neutral = 71.93, Negative = 4.28, TestNameBold = true, TestNameColor = "#808080" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMCLYDIN", Positive = 23.80, Neutral = 72.99, Negative = 3.21, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "RYDELTYX", Positive = 23.53, Neutral = 71.66, Negative = 4.81, TestNameBold = false, TestNameColor = "#FF0000", ProjectTemplateType="OverallImpressions" },
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "BLYTRAVE", Positive = 23.53, Neutral = 73.80, Negative = 2.67, TestNameBold = true, TestNameColor = "#808080" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMLIVLY", Positive = 23.26, Neutral = 74.33, Negative = 2.41, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "HAVUMFOR", Positive = 22.73, Neutral = 72.46, Negative = 4.81, TestNameBold = true, TestNameColor = "#808080" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMROBVI", Positive = 22.46, Neutral = 74.60, Negative = 2.94, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "REMPOJI", Positive = 22.46, Neutral = 75.67, Negative = 1.87, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "PAMTIVGO", Positive = 21.66, Neutral = 75.67, Negative = 2.67, TestNameBold = true, TestNameColor = "#808080" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "MIVZELTY", Positive = 21.39, Neutral = 75.94, Negative = 2.67, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMKELVIQ", Positive = 20.59, Neutral = 72.73, Negative = 6.68, TestNameBold = true, TestNameColor = "#009900" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMPZELTIV", Positive = 20.59, Neutral = 74.33, Negative = 5.08, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "EMCYZRA", Positive = 20.32, Neutral = 74.06, Negative = 5.61, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"},
                        new OverallImpressionsModel { ProjectName = "RACKEM", TestName = "HEVKESSI", Positive = 20.05, Neutral = 75.94, Negative = 4.01, TestNameBold = false, TestNameColor = "#FF0000" , ProjectTemplateType="OverallImpressions"}
                };

                await _context.OverallImpressionsTestTable.AddRangeAsync(data);
                await _context.SaveChangesAsync();
            }
        }
    }
}

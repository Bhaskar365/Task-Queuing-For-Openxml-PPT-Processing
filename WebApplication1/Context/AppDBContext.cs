
using Microsoft.EntityFrameworkCore;
using SharedModels;
using SharedModels.DTO;
using WebApplicationAPI.Models;

namespace WebApplication1.Context
{
    //public class AppDBContext : DbContext
    //{
    //    public AppDBContext(DbContextOptions options) : base(options) { }

    //    public DbSet<FitToConceptModel> FitToConceptTestTable { get; set; }

    //    public DbSet<OverallImpressionsModel> OverallImpressionsTestTable { get; set; }
    //}

    public partial class AppDBContext : DbContext
    {
        public AppDBContext()
        {
        }

        public AppDBContext(DbContextOptions options) : base(options) { }

        public DbSet<FitToConceptModel> FitToConceptTestTable { get; set; }

        public DbSet<OverallImpressionsModel> OverallImpressionsTestTable { get; set; }

        public  DbSet<Aev1> Aev1s { get; set; }

        public  DbSet<Aev2> Aev2s { get; set; }

        public  DbSet<Aev3> Aev3s { get; set; }

        public  DbSet<Memorability> Memorabilities { get; set; }

        public  DbSet<PersonalPreference> PersonalPreferences { get; set; }

        public  DbSet<Suffix> Suffixes { get; set; }

        public  DbSet<VerbalUnderstanding> VerbalUnderstandings { get; set; }

        public  DbSet<WrittenUnderstanding> WrittenUnderstandings { get; set; }

        public  DbSet<Likeability> Likeabilities { get; set; }

        public DbSet<Sala> Salas { get; set; }

        public DbSet<Sala154> Sala154 { get; set; }

        public DbSet<TaskLog> TaskLogs { get; set; }

        public DbSet<QTCModel> Qtc { get; set; }

        public DbSet<BrandexSafetyModel> BrandexSafety { get; set; }

        public DbSet<BrandexStrategicDistinctivenessModel> BrandexStrategicDistinctiveness { get; set; }

        public DbSet<MedicalTermsModel> MedicalTerms { get; set; }

        public DbSet<NonMedicalTermsModel> NonMedicalTerms { get; set; }

        public DbSet<ExaggModel> Untrue01 {get;set;}

        public DbSet<ExaggModel> Misleading02 { get; set; }

        public DbSet<ExaggModel> Exagg03 { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // attribute 1
            modelBuilder.Entity<Aev1>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("AEV1$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });

            // Logging Table
            modelBuilder.Entity<TaskLog>(entity =>
            {
                entity.HasNoKey().ToTable("TaskLoggingTable");
            });


            // QTC
            modelBuilder.Entity<QTCModel>(entity =>
            {
                entity.HasNoKey().ToTable("QTC$");
            });


            //Brandex Safety
            modelBuilder.Entity<BrandexSafetyModel>(entity =>
            {
                entity.HasNoKey().ToTable("BrandExSafety$");
            });

            //Brandex Strategic Distinctiveness
            modelBuilder.Entity<BrandexStrategicDistinctivenessModel>(entity =>
            {
                entity.HasNoKey().ToTable("BRANDEXSTRATEGICDISTINCTIVENESS$");
            });

            //medical terms
            modelBuilder.Entity<MedicalTermsModel>(entity =>
            {
                entity.HasNoKey().ToTable("MedicalTerms$");
            });

            //non-medical terms
            modelBuilder.Entity<NonMedicalTermsModel>(entity =>
            {
                entity.HasNoKey().ToTable("NonMedicalTerms$");
            });

            //untrue01
            modelBuilder.Entity<ExaggModel>(entity =>
            {
                entity.HasNoKey().ToTable("01 Untrue$");
            });

            //misleading
            modelBuilder.Entity<ExaggModel>(entity =>
            {
                entity.HasNoKey().ToTable("02 Misleading$");
            });

            //exagg
            modelBuilder.Entity<ExaggModel>(entity =>
            {
                entity.HasNoKey().ToTable("03 Exagg$");
            });


            modelBuilder.Entity<Aev2>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("AEV2$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });

            // sala proj:154
            modelBuilder.Entity<Sala154>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("Sheet1$");

                entity.Property(e => e.intBlue).HasColumnName("intBlue");
                entity.Property(e => e.intDrugNameCount).HasColumnName("intDrugNameCount");
                entity.Property(e => e.intGreen).HasColumnName("intGreen");
                entity.Property(e => e.intRed).HasColumnName("intRed");
                entity.Property(e => e.strDrugName0)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName0");
                entity.Property(e => e.strDrugName1)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName1");
                entity.Property(e => e.strDrugName10)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName10");
                entity.Property(e => e.strDrugName11)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName11");
                entity.Property(e => e.strDrugName12)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName12");
                entity.Property(e => e.strDrugName13)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName13");
                entity.Property(e => e.strDrugName14)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName14");
                entity.Property(e => e.strDrugName15)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName15");
                entity.Property(e => e.strDrugName16)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName16");
                entity.Property(e => e.strDrugName17)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName17");
                entity.Property(e => e.strDrugName18)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName18");
                entity.Property(e => e.strDrugName19)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName19");
                entity.Property(e => e.strDrugName2)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName2");
                entity.Property(e => e.strDrugName20)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName20");
                entity.Property(e => e.strDrugName21)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName21");
                entity.Property(e => e.strDrugName22)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName22");
                entity.Property(e => e.strDrugName23)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName23");
                entity.Property(e => e.strDrugName24)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName24");
                entity.Property(e => e.strDrugName25)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName25");
                entity.Property(e => e.strDrugName26)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName26");
                entity.Property(e => e.strDrugName27)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName27");
                entity.Property(e => e.strDrugName28)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName28");
                entity.Property(e => e.strDrugName29)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName29");
                entity.Property(e => e.strDrugName3)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName3");
                entity.Property(e => e.strDrugName4)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName4");
                entity.Property(e => e.strDrugName5)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName5");
                entity.Property(e => e.strDrugName6)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName6");
                entity.Property(e => e.strDrugName7)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName7");
                entity.Property(e => e.strDrugName8)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName8");
                entity.Property(e => e.strDrugName9)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName9");
                entity.Property(e => e.strDSIScore)
                    .HasMaxLength(255)
                    .HasColumnName("strDSIScore");
                entity.Property(e => e.strTestName)
                    .HasMaxLength(255)
                    .HasColumnName("strTestName");
                entity.Property(e => e.strTestNameTranslation)
                    .HasMaxLength(255)
                    .HasColumnName("strTestNameTranslation");
            });

            // Attribute Evaluation Aggregrate
            modelBuilder.Entity<Aev3>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("AEV3$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });


            // Memorability
            modelBuilder.Entity<Memorability>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("Memorability$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });


            // Personal Preference
            modelBuilder.Entity<PersonalPreference>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("PersonalPreference$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });

            // Suffix
            modelBuilder.Entity<Suffix>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("Suffix$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });

            // Verbal Understanding
            modelBuilder.Entity<VerbalUnderstanding>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("VerbalUnderstanding$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });


            // Written Understanding
            modelBuilder.Entity<WrittenUnderstanding>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("WrittenUnderstanding$");

                entity.Property(e => e.AverageColor).HasMaxLength(255);
                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
                entity.Property(e => e.ProjectTemplateType).HasMaxLength(255);
            });

            // Likeability
            modelBuilder.Entity<Likeability>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("Likeability$");

                entity.Property(e => e.ProjectName).HasMaxLength(255);
                entity.Property(e => e.TestName).HasMaxLength(255);
                entity.Property(e => e.TestNameColor).HasMaxLength(255);
            });

            // SALA
            modelBuilder.Entity<Sala>(entity =>
            {
                entity
                    .HasNoKey()
                    .ToTable("Sala$");

                entity.Property(e => e.intBlue).HasColumnName("intBlue");
                entity.Property(e => e.intDrugNameCount).HasColumnName("intDrugNameCount");
                entity.Property(e => e.intGreen).HasColumnName("intGreen");
                entity.Property(e => e.intRed).HasColumnName("intRed");
                entity.Property(e => e.strDrugName0)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName0");
                entity.Property(e => e.strDrugName1)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName1");
                entity.Property(e => e.strDrugName10)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName10");
                entity.Property(e => e.strDrugName11)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName11");
                entity.Property(e => e.strDrugName12)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName12");
                entity.Property(e => e.strDrugName13)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName13");
                entity.Property(e => e.strDrugName14)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName14");
                entity.Property(e => e.strDrugName15)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName15");
                entity.Property(e => e.strDrugName16)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName16");
                entity.Property(e => e.strDrugName17)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName17");
                entity.Property(e => e.strDrugName18)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName18");
                entity.Property(e => e.strDrugName19)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName19");
                entity.Property(e => e.strDrugName2)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName2");
                entity.Property(e => e.strDrugName20)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName20");
                entity.Property(e => e.strDrugName21)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName21");
                entity.Property(e => e.strDrugName22)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName22");
                entity.Property(e => e.strDrugName23)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName23");
                entity.Property(e => e.strDrugName24)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName24");
                entity.Property(e => e.strDrugName25)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName25");
                entity.Property(e => e.strDrugName26)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName26");
                entity.Property(e => e.strDrugName27)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName27");
                entity.Property(e => e.strDrugName28)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName28");
                entity.Property(e => e.strDrugName29)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName29");
                entity.Property(e => e.strDrugName3)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName3");
                entity.Property(e => e.strDrugName4)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName4");
                entity.Property(e => e.strDrugName5)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName5");
                entity.Property(e => e.strDrugName6)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName6");
                entity.Property(e => e.strDrugName7)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName7");
                entity.Property(e => e.strDrugName8)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName8");
                entity.Property(e => e.strDrugName9)
                    .HasMaxLength(255)
                    .HasColumnName("strDrugName9");
                entity.Property(e => e.strDSIScore)
                    .HasMaxLength(255)
                    .HasColumnName("strDSIScore");
                entity.Property(e => e.strTestName)
                    .HasMaxLength(255)
                    .HasColumnName("strTestName");
                entity.Property(e => e.strTestNameTranslation)
                    .HasMaxLength(255)
                    .HasColumnName("strTestNameTranslation");
            });


            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }

}


USE [OpenxmlCharts]
GO

SELECT 
		strTestName,
		strTestNameTranslation,
		dblAveragePage1
      ,dblPage1Weight
      ,dblAveragePage1Weighted
      ,dblAveragePage2
      ,dblPage2Weight
      ,dblAveragePage2Weighted
      ,dblAveragePage3
      ,dblPage3Weight
      ,dblAveragePage3Weighted
      ,dblAveragePage4
      ,dblPage4Weight
      ,dblAveragePage4Weighted
      ,dblAveragePage5
      ,dblPage5Weight
      ,dblAveragePage5Weighted
      ,dblAveragePage6
      ,dblPage6Weight
      ,dblAveragePage6Weighted
      ,dblAveragePage7
      ,dblPage7Weight
      ,dblAveragePage7Weighted
      ,dblAveragePage8
      ,dblPage8Weight
      ,dblAveragePage8Weighted
      ,dblIndex
      ,strDSIScore
      ,intRed
      ,intGreen
      ,intBlue
      ,CAST(boolBold AS BIT) AS boolBold
	INTO BRANDEXSTRATEGICDISTINCTIVENESS$_BACKUP
	from [dbo].[BRANDEXSTRATEGICDISTINCTIVENESS$]
	GO

	DROP TABLE [dbo].[BRANDEXSTRATEGICDISTINCTIVENESS$]

	CREATE TABLE [dbo].[BRANDEXSTRATEGICDISTINCTIVENESS$] (
		strTestName nvarchar(255) null,
		strTestNameTranslation nvarchar(255) null,
		dblAveragePage1 FLOAT NULL,
      dblPage1Weight  FLOAT NULL,
      dblAveragePage1Weighted FLOAT NULL,
      dblAveragePage2 FLOAT NULL,
      dblPage2Weight FLOAT NULL
      ,dblAveragePage2Weighted FLOAT NULL
      ,dblAveragePage3 FLOAT NULL
      ,dblPage3Weight FLOAT NULL
      ,dblAveragePage3Weighted FLOAT NULL
      ,dblAveragePage4 FLOAT NULL
      ,dblPage4Weight FLOAT NULL
      ,dblAveragePage4Weighted FLOAT NULL
      ,dblAveragePage5 FLOAT NULL
      ,dblPage5Weight FLOAT NULL
      ,dblAveragePage5Weighted FLOAT NULL
      ,dblAveragePage6 FLOAT NULL
      ,dblPage6Weight FLOAT NULL
      ,dblAveragePage6Weighted FLOAT NULL
      ,dblAveragePage7 FLOAT NULL 
      ,dblPage7Weight FLOAT NULL
      ,dblAveragePage7Weighted FLOAT NULL
      ,dblAveragePage8 FLOAT NULL
      ,dblPage8Weight FLOAT NULL
      ,dblAveragePage8Weighted FLOAT NULL
      ,dblIndex FLOAT NULL
      ,strDSIScore nvarchar(255) null
      ,intRed float null
      ,intGreen FLOAT NULL
      ,intBlue FLOAT NULL
      ,boolBold BIT null
	) ON [PRIMARY];
	GO

	INSERT INTO [dbo].[BRANDEXSTRATEGICDISTINCTIVENESS$]
	 (
		strTestName,
		strTestNameTranslation,
		dblAveragePage1
      ,dblPage1Weight
      ,dblAveragePage1Weighted
      ,dblAveragePage2
      ,dblPage2Weight
      ,dblAveragePage2Weighted
      ,dblAveragePage3
      ,dblPage3Weight
      ,dblAveragePage3Weighted
      ,dblAveragePage4
      ,dblPage4Weight
      ,dblAveragePage4Weighted
      ,dblAveragePage5
      ,dblPage5Weight
      ,dblAveragePage5Weighted
      ,dblAveragePage6
      ,dblPage6Weight
      ,dblAveragePage6Weighted
      ,dblAveragePage7
      ,dblPage7Weight
      ,dblAveragePage7Weighted
      ,dblAveragePage8
      ,dblPage8Weight
      ,dblAveragePage8Weighted
      ,dblIndex
      ,strDSIScore
      ,intRed
      ,intGreen
      ,intBlue,
	  boolBold
	 )
	 select * from BRANDEXSTRATEGICDISTINCTIVENESS$_BACKUP
	 GO

	 drop table BRANDEXSTRATEGICDISTINCTIVENESS$_BACKUP
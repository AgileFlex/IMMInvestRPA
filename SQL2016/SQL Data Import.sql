--sp_configure 'show advanced options', 1;  
--RECONFIGURE;
--GO 
--sp_configure 'Ad Hoc Distributed Queries', 1;  
--RECONFIGURE;  
--GO  

--EXEC master.[sys].[sp_MSset_oledb_prop] N'Microsoft.ACE.OLEDB.16.0', N'AllowInProcess', 1

USE IMMInvest

DECLARE @SQLText NVARCHAR(MAX) 
DECLARE @FileName VARCHAR(MAX) 

SET NOCOUNT ON;

TRUNCATE TABLE Requests;

INSERT INTO Requests 
(
	[Rue],
	[Data],
	[CuiCif],
	[Firma],
	[StareSolicitare]
)
--OPENROWSET
SELECT
	[Rue]					AS [Rue],
	[Dată înregistrare]		AS [Data],
	[CUI/CIF]				AS [CuiCif],
	[Firmă]					AS [Firma],
	[Stare solicitare]		AS [StareSolicitare]
FROM 
	OPENROWSET('Microsoft.ACE.OLEDB.16.0', 'Excel 12.0 Xml;Database=D:\UIPathRPA\ImmInvestRequests\Temp\imm-invest-requests-20-5-2020.xlsx', Sheet1$);

EXEC sp_getUniqueRequests
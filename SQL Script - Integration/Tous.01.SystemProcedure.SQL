IF EXISTS(SELECT * FROM sysobjects WHERE NAME='p_AddStockColumn' and xtype='p')
	DROP PROCEDURE [dbo].[p_AddStockColumn]
GO
CREATE PROCEDURE [dbo].[p_AddStockColumn]
    	@FieldName		NVARCHAR(100),
	@FieldType		INT,
	@HeadOrEntry	INT,
	@IsNull			INT = 0
AS
    SET NOCOUNT ON
	
	DECLARE @sSQL	NVARCHAR(4000)
	DECLARE @TableName	NVARCHAR(100)
	DECLARE @TypeName	NVARCHAR(100)
	DECLARE @Where		NVARCHAR(1000)

	IF @HeadOrEntry <= 0
		SET @TableName = 'ICStockBill'
	ELSE
		SET @TableName = 'ICStockBillEntry'

	IF @FieldType = 0
	BEGIN
		SET @TypeName = ' INT'
		IF @IsNull = 0
			SET @Where = ' NOT NULL DEFAULT(0)'
		ELSE
			SET @Where = ' NULL DEFAULT(0)'
	END

	IF @FieldType = 1
	BEGIN
		SET @TypeName = ' NVARCHAR(255)'
		IF @IsNull = 0
			SET @Where = ' NOT NULL DEFAULT('''')'
		ELSE
			SET @Where = ' NULL DEFAULT('''')'
	END

	IF @FieldType = 2
	BEGIN
		SET @TypeName = ' DATETIME '
		SET @Where = ' NULL'
	END

	IF @FieldType = 3
	BEGIN
		SET @TypeName = ' DECIMAL(28,10)'
		SET @Where = ' NOT NULL DEFAULT(0)'
	END

	IF @FieldType = 4
	BEGIN
		SET @TypeName = ' BIT'
		SET @Where = ' NOT NULL DEFAULT(0)'
	END

	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+''')) '
				+ ' ALTER TABLE '+@TableName+' ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_1'')) '
				+ ' ALTER TABLE '+@TableName+'_1 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_2'')) '
				+ ' ALTER TABLE '+@TableName+'_2 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_5'')) '
				+ ' ALTER TABLE '+@TableName+'_5 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_10'')) '
				+ ' ALTER TABLE '+@TableName+'_10 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_21'')) '
				+ ' ALTER TABLE '+@TableName+'_21 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET  @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_24'')) '
				+ ' ALTER TABLE '+@TableName+'_24 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_28'')) '
				+ ' ALTER TABLE '+@TableName+'_28 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_29'')) '
				+ ' ALTER TABLE '+@TableName+'_29 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+'_41'')) '
				+ ' ALTER TABLE '+@TableName+'_41 ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
GO

IF EXISTS(SELECT * FROM sysobjects WHERE NAME='p_AddTableColumn' and xtype='p')
	DROP PROCEDURE [dbo].[p_AddTableColumn]
GO
CREATE PROCEDURE [dbo].[p_AddTableColumn]
	@TableName		NVARCHAR(100),
    @FieldName		NVARCHAR(100),
	@FieldType		INT,
	@IsNull			INT = 0
AS
    SET NOCOUNT ON
	
	DECLARE @sSQL	NVARCHAR(4000)
	DECLARE @TypeName	NVARCHAR(100)
	DECLARE @Where		NVARCHAR(1000)

	IF @FieldType = 0
	BEGIN
		SET @TypeName = ' INT'
		IF @IsNull = 0
			SET @Where = ' NOT NULL DEFAULT(0)'
		ELSE
			SET @Where = ' NULL DEFAULT(0)'
	END

	IF @FieldType = 1
	BEGIN
		SET @TypeName = ' NVARCHAR(255)'
		IF @IsNull = 0
			SET @Where = ' NOT NULL DEFAULT('''')'
		ELSE
			SET @Where = ' NULL DEFAULT('''')'
	END

	IF @FieldType = 2
	BEGIN
		SET @TypeName = ' DATETIME '
		SET @Where = ' NULL'
	END

	IF @FieldType = 3
	BEGIN
		SET @TypeName = ' DECIMAL(28,10)'
		SET @Where = ' NOT NULL DEFAULT(0)'
	END

	IF @FieldType = 4
	BEGIN
		SET @TypeName = ' BIT'
		SET @Where = ' NOT NULL DEFAULT(0)'
	END

	SET @sSQL = 'IF NOT EXISTS(SELECT 1 FROM Syscolumns WHERE NAME='''+@FieldName+'''  AND ID=object_id('''+@TableName+''')) '
				+ ' ALTER TABLE '+@TableName+' ADD '+@FieldName+@TypeName+@Where
	EXEC(@sSQL)
GO

IF EXISTS(SELECT * FROM sysobjects WHERE NAME='p_AddNewComponent' and xtype='p')
	DROP PROCEDURE [dbo].[p_AddNewComponent]
GO
CREATE PROCEDURE [dbo].[p_AddNewComponent]  
    @ClassID	NVARCHAR(100),  
	@Event		NVARCHAR(100),  
	@Com		NVARCHAR(100),  
	@First		BIT = 0  
AS 
    SET NOCOUNT ON  
  
	DECLARE @Component NVARCHAR(2000)  
	DECLARE @ComEx NVARCHAR(2000)  
	DECLARE @ComSub NVARCHAR(2000)  
	DECLARE @Index1 INT  
	DECLARE @Index2 INT  

	SELECT @Component = FComponentExt FROM ICClassType WHERE FID = @ClassID  
	SELECT @Index1 = Charindex(@Event,@Component,1)  
	SELECT @Index2 = Charindex('|',@Component,@Index1)  
	SELECT @ComSub = SubString(@Component,@Index1,@Index2-@Index1)  

	IF Charindex(@Com,@ComSub,1)=0  
	BEGIN  
		SELECT @ComEx = LEFT(@Component,@Index2-1)  
		IF RIGHT(@ComEx,1) = ';' OR RIGHT(@ComEx,1) = '='  
			SELECT @Component = @ComEx + @Com + RIGHT(@Component,LEN(@Component)+1-@Index2)  
		ELSE  
			SELECT @Component = @ComEx + ';' + @Com + RIGHT(@Component,LEN(@Component)+1-@Index2)  
		UPDATE ICClassType SET FComponentExt = @Component WHERE FID = @ClassID  
	END
GO
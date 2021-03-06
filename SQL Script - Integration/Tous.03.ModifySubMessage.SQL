------------------------------------------------------------
--Craete Date:2018.09.15
--Last Modify:
--Description:辅助资料
------------------------------------------------------------
--************ 新增辅助资料Purchase Type，用于区分采购类型 **************--
DELETE t_SubMesType  WHERE FName='Purchase Type' AND FTypeID = '11100'
INSERT INTO t_SubMesType([FTypeID],[FDetail],[FName],[FType],[FGRType])
VALUES(11100, 0, 'Purchase Type', 1, 0)

--新增辅助资料Store Type，辅助资料详细--
DELETE t_SubMessage  WHERE FTypeID = '11100'
INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11101, 'MERC', 11100, 'Merchandise', 11100, 0, 0, 1)

INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11102, 'NONS', 11100, 'Non-Sale Material', 11100, 0, 0, 1)


--**************  新增辅助资料Inventory Type，用于区分仓库类型  **************--
DELETE t_SubMesType  WHERE FName='Inventory Type' AND FTypeID = '11200'
INSERT INTO t_SubMesType([FTypeID],[FDetail],[FName],[FType],[FGRType])
VALUES(11200, 0, 'Inventory Type', 1, 0)

--新增辅助资料Inventory Type，用于区分仓库类型,辅助资料详细--
DELETE t_SubMessage  WHERE FTypeID = '11200'
INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11201, 'Store', 11200, 'Store', 11200, 0, 0, 1)

INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11202, 'Office', 11200, 'Office', 11200, 0, 0, 1)

INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11203, 'InTran', 11200, 'In-transit', 11200, 0, 0, 1)
	
INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11204, 'WH', 11200, 'Warehouse', 11200, 0, 0, 1)


--**************  新增辅助资料Adjustment Reason，用于记录调整原因  **************--
DELETE t_SubMesType  WHERE FName='Adjustment Reason' AND FTypeID = '11400'
INSERT INTO t_SubMesType([FTypeID],[FDetail],[FName],[FType],[FGRType])
VALUES(11400, 0, 'Adjustment Reason', 1, 0)

--新增辅助资料Adjustment Reason，用于记录调整原因,辅助资料详细--
DELETE t_SubMessage  WHERE FTypeID = '11400'
INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11401, 'Inv', 11400, 'Inventory Gain/Loss', 11400, 0, 0, 1)

INSERT INTO t_SubMessage([FBrNo],[FInterID],[FID],[FParentID],[FName],[FTypeID],[FDetail],[FDeleted],[FSystemType])
VALUES(0, 11402, 'MER', 11400, 'Merchandising', 11400, 0, 0, 1)

--**************  新增辅助资料经销商清单  **************--
DELETE t_SubMesType  WHERE FName='经销商清单' AND FTypeID = '11500'
INSERT INTO t_SubMesType([FTypeID],[FDetail],[FName],[FType],[FGRType])
VALUES(11500, 0, '经销商清单', 1, 0)



----------因此视图涉及到t_icItem表，所以视图放在此处-----
--IF EXISTS(SELECT 1 FROM SysObjects WHERE NAME = 'v_K3View21' AND XTYPE = 'V')
--	DROP VIEW v_K3View21
--GO
--CREATE VIEW v_K3View21
--AS
--SELECT v1.FlowType,v1.FInterID,isnull(v1.FIsExport,0)as FIsExport,v1.FTranType,v1.FCheckerID,56 AS RECORDTYPE,''AS DOCUMENTTYPE,''AS DOCUMENTLINETYPE,'N' AS AUTORECEIVEFLAG,t2.FNumber as STORECODE,
--t1.FNumber as DELIVERYLOCATION,t3.FStyleNumber as PRODUCTCODE,'' as PRODUCTREFERENCE,t3.FColorEx as COLORCODE,T3.FSizeEx as SIZECODE,
--'' as SKU,u1.FAuxQty as QUANTITY,v1.FReference as STOREORDERNUMBER,'' as VENDORORDERNUMBER,v1.FReference as DOCUMENTNUMBER,
--'' as VENDORSHIPMENTNUMBER,''as CURRENCY,'' as VENDORCODE,'' as PURCHASEPRICE,'' as DISCOUNT,''as LANDEDCOST,''as TAXCOST,
--'' as AVERAGECOST,v1.FDATE as SHIPMENTDATE,v1.FDATE  as DELIVERYDATE,v1.FDATE  as ORDERDATE,'' as REQUESTEDDELIVERYDATE,
--'' as CANCELLATIONDATE,'' as BRAND,'' as SERIALNUMBER,'' as PRODUCTLINE,'' as CUSTOMERORDER,'Y' as ASSIGNTOBIN,
--'' as BINCATEGORY,'' as QTYDECIMAL,'' as UCC128,t3.FNumber as EXTENDEDEXTERNALSKU,'' as QTYSIGN,'' as BINCODE,
--'' as CARRIERCODE,'' as CARRIERTRACKINGNUMBER,'' as RETAILPRICE,'' as CARRIERAMOUNT,'' as TRANSFERSENTDESTINATIONBINTYPE,
--'' as ADJQTY,'' as SUBPONUMBER,'DEFR' as ReasonCode
--FROM ICStockBill v1 
--INNER JOIN ICStockBillEntry u1 on v1.FInterID=u1.FInterID
--INNER JOIN T_STOCK T1 ON T1.FItemID=u1.FDestinationExID --调入仓(需取新增字段仓库的值待改 )
--INNER JOIN T_STOCK T2 ON T2.FItemID=u1.FSCStockID --调出仓
--INNER JOIN T_ICItem T3 ON T3.FItemID=u1.FItemID  --物料
--LEFT JOIN T_SUBMESSAGE T4 ON T4.FINTERID=T3.FColorExID --Color
----left join t_TB_StockDeliveryNotice t5 on t5.FBillNo=v1.FReference ---关联stocknotice 取SSCC
----left join t_TB_StockDeliveryContainer t6 on t6.FID=t5.FID 


--GO
--IF EXISTS(SELECT 1 FROM SysObjects WHERE NAME = 'v_K3View19' AND XTYPE = 'V')
--	DROP VIEW v_K3View19
--GO
--CREATE VIEW v_K3View19
--AS
--SELECT v1.FlowType,v1.FInterID,isnull(v1.FIsExport,0)as FIsExport,v1.FTranType,v1.FCheckerID,56 AS RECORDTYPE,''AS DOCUMENTTYPE,''AS DOCUMENTLINETYPE,'N' AS AUTORECEIVEFLAG,t2.FNumber as STORECODE,
--right(t1.FNumber,4) as DELIVERYLOCATION,t3.FStyleNumber as PRODUCTCODE,'' as PRODUCTREFERENCE,T4.FID as COLORCODE, T3.FSizeEx as SIZECODE,
--'' as SKU,u1.FAuxQty as QUANTITY,v1.FReference as STOREORDERNUMBER,'' as VENDORORDERNUMBER,v1.FReference as DOCUMENTNUMBER,
--'' as VENDORSHIPMENTNUMBER,''as CURRENCY,'' as VENDORCODE,'1' as PURCHASEPRICE,'' as DISCOUNT,'1' as LANDEDCOST,''as TAXCOST,
--'1' as AVERAGECOST,CONVERT(varchar(100), v1.FDATE, 112) as SHIPMENTDATE,CONVERT(varchar(100), v1.FDATE, 112)   as DELIVERYDATE,CONVERT(varchar(100), v1.FDATE, 112)  as ORDERDATE,'' as REQUESTEDDELIVERYDATE,
--'' as CANCELLATIONDATE,'' as BRAND,'' as SERIALNUMBER,'' as PRODUCTLINE,'' as CUSTOMERORDER,'Y' as ASSIGNTOBIN,
--'' as BINCATEGORY,'' as QTYDECIMAL,u1.FContainerNumber  as UCC128,t3.FNumber as EXTENDEDEXTERNALSKU,'' as QTYSIGN,'' as BINCODE,
--'' as CARRIERCODE,'' as CARRIERTRACKINGNUMBER,'' as RETAILPRICE,'' as CARRIERAMOUNT,'' as TRANSFERSENTDESTINATIONBINTYPE,
--'' as ADJQTY,'' as SUBPONUMBER,'DEFR' as ReasonCode
--FROM ICStockBill v1 
--INNER JOIN ICStockBillEntry u1 on v1.FInterID=u1.FInterID
--INNER JOIN T_STOCK T1 ON T1.FItemID=u1.FDestinationExID --调入仓(需取新增字段仓库的值待改 )
--INNER JOIN T_STOCK T2 ON T2.FItemID=u1.FSCStockID --调出仓
--INNER JOIN T_ICItem T3 ON T3.FItemID=u1.FItemID  --物料
--LEFT JOIN T_SUBMESSAGE T4 ON T4.FINTERID=T3.FColorExID --Color
--GO
-------------------------------------------------

SELECT G.*,
A.[Files-XOML],A.[Files-XSN],A.[Files-CSS],A.[Files-JS],A.[Files-JQuery],A.[Files-Angular],A.[Files-Bootstrap],
B.[Lists-Over5KItems],B.[Lists-Unthrottled],B.[Lists-NumInboundEmail],B.[Lists-LastModifiedInboundEmail],B.[Lists-LastModified],
C.[Feature-PublishingSite],C.[Feature-MinimalPublishingSite],
D.[Feature-PublishingWeb],D.[Feature-MinimalPublishingWeb],
E.[Alerts-Immed],
F.[Alerts-Sched]
FROM 
(
	SELECT Webs.Id,
	SUM(CASE WHEN ((AllDocs.ExtensionForFile = 'xoml') AND (AllDocs.DirName NOT LIKE '%_catalogs%')) THEN 1 ELSE 0 END) AS 'Files-XOML',
	SUM(CASE WHEN ((AllDocs.ExtensionForFile = 'xsn') AND (AllDocs.DirName NOT LIKE '%_catalogs%')) THEN 1 ELSE 0 END) AS 'Files-XSN',
	SUM(CASE WHEN ((AllDocs.ExtensionForFile = 'css') AND (AllDocs.DirName NOT LIKE '%_catalogs%') AND (AllDocs.DirName NOT LIKE '%en-us%Core Styles%')) THEN 1 ELSE 0 END) AS 'Files-CSS',
	SUM(CASE WHEN ((AllDocs.ExtensionForFile = 'js') AND (AllDocs.DirName NOT LIKE '%_catalogs%')) THEN 1 ELSE 0 END) AS 'Files-JS',
	SUM(CASE WHEN ((AllDocs.LeafName LIKE '%jquery%js') AND (AllDocs.DirName NOT LIKE '%_catalogs%')) THEN 1 ELSE 0 END) AS 'Files-JQuery',
	SUM(CASE WHEN ((AllDocs.LeafName LIKE '%angular%js') AND (AllDocs.DirName NOT LIKE '%_catalogs%')) THEN 1 ELSE 0 END) AS 'Files-Angular',
	SUM(CASE WHEN ((AllDocs.LeafName LIKE '%bootstrap%css') AND (AllDocs.DirName NOT LIKE '%_catalogs%')) THEN 1 ELSE 0 END) AS 'Files-Bootstrap'
	FROM Webs
	LEFT JOIN AllDocs
	ON Webs.Id = AllDocs.WebId
	GROUP BY Webs.Id
) AS A
INNER JOIN
(
	SELECT Webs.Id,
	SUM(CASE WHEN AllListsAux.ItemCount > 5000 THEN 1 ELSE 0 END) AS 'Lists-Over5KItems',
	SUM(CASE WHEN AllLists.tp_NoThrottleListOperations <> 0 THEN 1 ELSE 0 END) AS 'Lists-Unthrottled',
	SUM(CASE WHEN AllLists.tp_EmailAlias IS NOT NULL THEN 1 ELSE 0 END) AS 'Lists-NumInboundEmail',
	MAX(CASE WHEN AllLists.tp_EmailAlias IS NOT NULL THEN AllListsAux.Modified ELSE 0 END) AS 'Lists-LastModifiedInboundEmail',
	MAX(AllListsAux.Modified) AS 'Lists-LastModified'
	FROM            Webs
	LEFT JOIN AllLists
	ON Webs.Id = AllLists.tp_WebId
	LEFT JOIN AllListsAux
	ON AllLists.tp_Id = AllListsAux.ListId
	GROUP BY Webs.Id
) AS B
ON A.Id=B.Id
INNER JOIN
(
	SELECT Webs.Id,
	SUM(CASE WHEN F1.FeatureId = 'F6924D36-2FA8-4F0B-B16D-06B7250180FA' THEN 1 ELSE 0 END) AS 'Feature-PublishingSite',
	SUM(CASE WHEN F1.FeatureId = '63FDC6AC-DBB4-4247-B46E-A091AEFC866F' THEN 1 ELSE 0 END) AS 'Feature-MinimalPublishingSite'
	FROM            Webs 
	LEFT JOIN Features AS F1
	ON Webs.SiteId = F1.SiteId
	GROUP BY Webs.Id
) AS C
ON A.Id=C.Id
INNER JOIN
(
	SELECT Webs.Id,
	SUM(CASE WHEN F2.FeatureId = '94C94CA6-B32F-4DA9-A9E3-1F3D343D7ECB' THEN 1 ELSE 0 END) AS 'Feature-PublishingWeb',
	SUM(CASE WHEN F2.FeatureId = 'A4A489B1-5420-40C3-8DB7-247C9FC51CA9' THEN 1 ELSE 0 END) AS 'Feature-MinimalPublishingWeb'
	FROM            Webs 
	LEFT JOIN Features AS F2
	ON Webs.Id = F2.WebId
	GROUP BY Webs.Id
) AS D
ON A.Id=D.Id
INNER JOIN
(
	SELECT Webs.Id, 
	SUM(CASE WHEN ImmedSubscriptions.Id IS NULL THEN 0 ELSE 1 END) AS 'Alerts-Immed'
	FROM Webs
	LEFT JOIN ImmedSubscriptions
	ON Webs.Id = ImmedSubscriptions.WebId
	GROUP BY Webs.Id
) AS E
ON A.Id=E.Id
INNER JOIN
(
	SELECT Webs.Id, 
	SUM(CASE WHEN SchedSubscriptions.Id IS NULL THEN 0 ELSE 1 END) AS 'Alerts-Sched'
	FROM Webs
	LEFT JOIN SchedSubscriptions
	ON Webs.Id = SchedSubscriptions.WebId
	GROUP BY Webs.Id
) AS F
ON A.Id=F.Id
INNER JOIN
(
	SELECT Id,FullUrl,title,RequestAccessEmail,WebTemplate,AlternateCSSUrl,CustomJSUrl,MasterUrl,CustomMasterUrl 
	FROM Webs
) AS G
ON A.Id=G.Id
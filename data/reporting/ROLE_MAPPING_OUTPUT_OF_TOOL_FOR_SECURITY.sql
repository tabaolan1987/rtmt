SELECT 
	[NTID],
	[fname] AS [First Name],
	[lname] AS [Last Name],
	[jobTitle] AS [Job Title],
	'' AS [Performace Unit],
	[purchasingOrg] AS [Purchasing Org],
	[jobTitle] AS [Job Title 2],
	[siteLocation] AS[Location-Site],
	'' AS [HR Status],
	'' AS [Requisition/Indent ($Limit in Thousands)],
	'' AS [Invoice Approval ($Limit in Thousands)],
	'' AS [3rd Party Commitment ($Limit in Thousands)]
FROM user_data ORDER BY [ntid]
Where user_data.deleted=0
=====
SELECT 
	{% 
		SRM Lead Requestor,Backbone Reviewer/Approver,
		Procurement Display & Reporting,
		Contract Display & Reporting,Service Entry Sheet Creator,
		Operational Buyer,Contract Owner,Procurement Catalogue Processor,
		PSCM Specialist,PSCM Team Leader,Sourcing Coordinator,
		T&C Supressor,AVL Suppressor,Confidential Contracts Management,
		PSCM Approver,Bidder Creator,Goods Receiver,
		Title Transfer Receiver,Warehouse / Logistics Specialist,
		MRP Specialist,Inventory Scrapper,Inventory Transfer Specialist,
		Stock Count Administrator,Stock Count Variance Processor,
		Inventory Requester,Materials Fabrication Requestor,
		MM Financial Approver,Materials Management Display & Reporting,
		CRP Processor,Cost Allocation Administrator,
		Shipment Specialist,Rental Specialist,Materials Expediter,
		POQR Library Administrator,POQR Approver,POQR Document Reviewer,
		Inventory Optimisation Analyst,WM Supervisor,WM Scheduler,
		WM Planner,WM Advanced Planner 1,WM Advanced Planner 2,
		WM Technician,WM Non-Planning Technician,WM Mobile Technician,
		WM Microsoft Project (MSP),WM Senior Leadership,
		WM Display & Reporting,AP Invoice Processor,
		WM Occasional Work Request Creator,
		WM Local Work Management Administrator,
		WM Regional Work Management Administrator,
		WM Regional Maximo Labour Data Steward,Master Data Administrator - Item,
		Master Data Steward - MDM Team,Master Data Administrator - Warehouse,
		Master Data Administrator - BOM/Product Structure,
		Master Data Administrator - Service,Master Data Administrator - PSCM,
		Cost Approver Maintainer,Material Master Technical Approver,
		AVL Co-ordinator,Vendor Maintainer - SQM,Vendor Maintainer - HSSE,
		Tax Maintainer,Order Settlement,Accruals Co-ordinator,Tax Expert,
		Finance Integration Display and Reporting,Accounting Object Analyst,
		Invoice Exception and Workflow Analyst,MI Query Writer,Regional Backbone Administrator
		| 
		(select top 1 IIF(NOT ISNULL(bpRole.BpRoleStandardName), "Y","") 
from (((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.id)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.id = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
where udata.ntid = UD.ntid
and udata.Deleted = 0 and sp.Deleted = 0
and spAc.Deleted = 0 and ac.Deleted = 0
and AcBpMapp.Deleted = 0 and bpRole.Deleted = 0
and bpRole.BpRoleStandardName = '(%VALUE%)'
and spAc.function_region='(%RG_F_ID%)')
				AS [(%VALUE%)]
			%}
FROM user_data AS UD 
where UD.region = '(%RG_NAME%)'
ORDER BY UD.ntid


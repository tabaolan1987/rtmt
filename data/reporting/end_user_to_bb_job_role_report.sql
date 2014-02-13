SELECT 
	[NTID],
	[GPID],
	[fname] AS [First Name],
	[lname] AS [Last Name],
	[email] AS [E-mail address],
	[omsSubfunction] AS [Function (OMS)/ Sub-function],
	[departmentBusiness] AS [Department or Business Unit],
	[Specialism],
	[jobTitle] AS [Job Title],
	[sponsorForeName] AS [Line Manager/ Sponsor Forename],
	[sponsorSurname] AS [Line Manager/ Sponsor Surname],
	[VTA],
	[Country],
	[contractor] AS [Contractor?(Y/N)],
	[SFunction] AS [Standard Function],
	[SdSubFunction] AS [Standard Sub Function],
	[STeam] AS [Standard Team],
	[Spare1],
	[Spare2],
	[Spare3],
	[Spare4],
	[Spare5],
	[Spare6],
	[Spare7],
	[Spare8],
	[Spare9],
	[Spare10],
	[Spare11],
	[Spare12],
	[Spare13],
	[Spare14],
	[Spare15],
	[Spare16],
	[Spare17],
	[Spare18],
	[Spare19],
	[Spare20]
FROM user_data
Where user_data.deleted=0
and user_data.SFunction='(%RG_F_NAME%)'
ORDER BY [ntid]
=====
SELECT 
	{% 
		SRM Lead Requester,
		Standard Desktop Confirmation Requester,
		Central Desktop Confirmation Requester,Backbone Reviewer/Approver,
		Procurement Display & Reporting,
		Contract Display & Reporting,
		Service Entry Sheet Creator,
		Operational Buyer,Contract Owner,
		Procurement Catalogue Processor,
		PSCM Specialist,PSCM Team Leader,
		Sourcing Coordinator,
		T&C Suppressor,Confidential Contracts Management,
		PSCM Approver,
		Bidder Creator,
		Goods Receiver,
		Title Transfer Receiver,
		Warehouse / Logistics Specialist,MRP Specialist,
		Strategic Materials Planner,Demand Planning MRP Specialist,
		VMI Administrator,
		Third Party Inventory Administrator,
		Inventory Scrapper,
		Inventory Transfer Specialist,
		Stock Count Administrator,
		Stock Count Variance Processor,
		Inventory Requestor,
		Materials Fabrication Requestor,
		MM Financial Approver,
		Materials Management Display & Reporting ,CRP Processor,
		Cost Allocation Administrator (GoM),
		Shipment Specialist,
		Rental Specialist,
		Materials Expediter,
		Inventory Optimization Analyst,
		Inventory Optimisation Analyst,
		POQR Library Administrator,
		POQR Document Reviewer,
		POQR Approver,
		WM Supervisor,
		WM Scheduler ,
		WM Planner,
		WM Advanced Planner 1,
		WM Advanced Planner 2,
		WM Technician,
		WM Mobile Technician,
		WM Microsoft Project,
		WM Senior Leadership,
		WM Display & Reporting,
		WM Local Work Management Administrator,
		WM Regional Work Management Administrator,
		Master Data Administrator - Item,Master Data Administrator - Warehouse,
		Master Data Administrator - BOM/Product Structure,
		Master Data Administrator - Service,
		MDM - Global Data Steward,
		MDM - Local Data Steward,
		MDM - Display,
		Master Data Administrator - PSCM,
		Cost Approver Maintainer,
		Vendor Data Requestor (Egypt only),
		Vendor Maintainer - HSSE,GWO Data Maintainer,
		GWO Data Display,
		Order Settlement,
		Invoice Exception and Workflow Analyst,
		Tax Maintainer,
		Accounting Object Analyst,
		Tax Expert,
		AP Invoice Processor (GFT Job Role),
		Finance Integration Display and Reporting,
		MI Query Writer,
		Regional Maximo Labor Data Steward,
		Regional Backbone Administrator
		| 
		(select top 1 IIF(NOT ISNULL(bpRole.BpRoleStandardName), "Y","") 
from (user_data_mapping_role as UMR 
inner join BpRoleStandard as bpRole
on UMR.idBpRoleStandard = bpRole.id)
where UMR.idUserdata = UD.ntid and UMR.Deleted=0
and UD.deleted = 0
and bpRole.Deleted = 0
and bpRole.BpRoleStandardName = '(%VALUE%)'
and UMR.idFunction='(%RG_F_ID%)')
				AS [(%VALUE%)]
			%}
FROM user_data AS UD
where UD.SFunction='(%RG_F_NAME%)'
ORDER BY UD.ntid

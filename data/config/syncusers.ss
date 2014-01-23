# IMPORTANT
ntid : NTID
gpid : GPID
fname : First Name
lname : Last Name
email : E-mail address
omsSubfunction : Function (OMS)/ Sub-function
departmentBusiness : Department or Business Unit
specialism : Specialism
jobTitle : Job Title
sponsorForeName : Line Manager/ Sponsor Forename
sponsorSurname : Line Manager/ Sponsor Surname
vta : VTA
country : Country
contractor : Contractor?
changeNetworkLevel : Change Network Level
dofa : DofA (Y/N)
dofaType : DofA Type  (Indent, Commitment, Sourcing)
region : Region
siteLocation : Maximo Site Location
purchasingOrg : Purchasing Org
spare1 : Spare1
spare2 : Spare2
spare3 : Spare3
spare4 : Spare4
spare5 : Spare5
spare6 : Spare6
spare7 : Spare7
spare8 : Spare8
spare9 : Spare9
spare10 : Spare10
spare11 : Spare11
spare12 : Spare12
spare13 : Spare13
spare14 : Spare14
spare15 : Spare15
spare16 : Spare16
spare17 : Spare17
spare18 : Spare18
spare19 : Spare19
spare20 : Spare20
# Custom import format:   "insert query" | "check value" : "list of column to check"
# INSERT INTO user_data_mapping_role(idUserdata, idRegion, idBpRoleStandard) SELECT top 1 [ntid],idRegion,idBpRoleStandard FROM BpRoleStandard, Region where BpRoleStandardName = [value] And RegionName = [region_name] | y : SRM Lead Requester,Standard Desktop Confirmation Requester,Central Desktop Confirmation Requester,Backbone Reviewer/Approver,Procurement Display & Reporting,Contract Display & Reporting,Service Entry Sheet Creator,Operational Buyer,Contract Owner,Procurement Catalogue Processor,PSCM Specialist,PSCM Team Leader,Sourcing Coordinator,T&C Suppressor,Confidential Contracts Management,PSCM Approver,Bidder Creator,Goods Receiver,Title Transfer Receiver,Warehouse / Logistics Specialist,MRP Specialist,Strategic Materials Planner,Demand Planning MRP Specialist,VMI Administrator,Third Party Inventory Administrator,Inventory Scrapper,Inventory Transfer Specialist,Stock Count Administrator,Stock Count Variance Processor,Inventory Requestor,Materials Fabrication Requestor,MM Financial Approver,Materials Management Display & Reporting,CRP Processor,Cost Allocation Administrator (GoM),Shipment Specialist,Rental Specialist,Materials Expediter,Inventory Optimisation Analyst,POQR Library Administrator,POQR Document Reviewer,POQR Approver,WM Supervisor,WM Scheduler,WM Planner,WM Advanced Planner 1,WM Advanced Planner 2,WM Technician,WM Mobile Technician,WM Microsoft Project,WM Senior Leadership,WM Display & Reporting,WM Local Work Management Administrator,WM Regional Work Management Administrator,Master Data Administrator - Item,Master Data Administrator - Warehouse,Master Data Administrator - BOM/Product Structure,Master Data Administrator - Service,MDM - Global Data Steward,MDM - Local Data Steward,MDM - Display,Master Data Administrator - PSCM,Cost Approver Maintainer,Vendor Data Requestor (Egypt only),Vendor Maintainer - HSSE,GWO Data Maintainer,GWO Data Display,Order Settlement,Invoice Exception and Workflow Analyst,Tax Maintainer,Accounting Object Analyst,Tax Expert,AP Invoice Processor (GFT Job Role),Finance Integration Display and Reporting,MI Query Writer,Regional Maximo Labor Data Steward,Regional Backbone Administrator
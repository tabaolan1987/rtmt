# IMPORTANT
NTID : ntid
GPID : gpid
First Name : fname
Last Name : lname
E-mail address : email
Function (OMS)/ Sub-function : omsSubfunction
Department or Business Unit : departmentBusiness
Specialism : specialism
Job Title : jobTitle
Line Manager/ Sponsor Forename : sponsorForeName
Line Manager/ Sponsor Surname : sponsorSurname
VTA : vta
Country : country
Contractor? : contractor
Change Network Level : changeNetworkLevel
DofA (Y/N) : dofa
DofA Type  (Indent, Commitment, Sourcing) : dofaType
Region : region
Spare1 : spare1
Spare2 : spare2
Spare3 : spare3
Spare4 : spare4
Spare5 : spare5
#Maximo Site Location : siteLocation
#Purchasing Org : purchasOrg
#Optional Field 1 : Optionalfield1
#Optional Field 2 : Optionalfield2
# Custom import format:   "insert query" | "check value" : "list of column to check"
# INSERT INTO user_data_mapping_role(idUserdata, idRegion, idBpRoleStandard) SELECT top 1 [ntid],idRegion,idBpRoleStandard FROM BpRoleStandard, Region where BpRoleStandardName = [value] And RegionName = [region_name] | y : SRM Lead Requester,Standard Desktop Confirmation Requester,Central Desktop Confirmation Requester,Backbone Reviewer/Approver,Procurement Display & Reporting,Contract Display & Reporting,Service Entry Sheet Creator,Operational Buyer,Contract Owner,Procurement Catalogue Processor,PSCM Specialist,PSCM Team Leader,Sourcing Coordinator,T&C Suppressor,Confidential Contracts Management,PSCM Approver,Bidder Creator,Goods Receiver,Title Transfer Receiver,Warehouse / Logistics Specialist,MRP Specialist,Strategic Materials Planner,Demand Planning MRP Specialist,VMI Administrator,Third Party Inventory Administrator,Inventory Scrapper,Inventory Transfer Specialist,Stock Count Administrator,Stock Count Variance Processor,Inventory Requestor,Materials Fabrication Requestor,MM Financial Approver,Materials Management Display & Reporting,CRP Processor,Cost Allocation Administrator (GoM),Shipment Specialist,Rental Specialist,Materials Expediter,Inventory Optimisation Analyst,POQR Library Administrator,POQR Document Reviewer,POQR Approver,WM Supervisor,WM Scheduler,WM Planner,WM Advanced Planner 1,WM Advanced Planner 2,WM Technician,WM Mobile Technician,WM Microsoft Project,WM Senior Leadership,WM Display & Reporting,WM Local Work Management Administrator,WM Regional Work Management Administrator,Master Data Administrator - Item,Master Data Administrator - Warehouse,Master Data Administrator - BOM/Product Structure,Master Data Administrator - Service,MDM - Global Data Steward,MDM - Local Data Steward,MDM - Display,Master Data Administrator - PSCM,Cost Approver Maintainer,Vendor Data Requestor (Egypt only),Vendor Maintainer - HSSE,GWO Data Maintainer,GWO Data Display,Order Settlement,Invoice Exception and Workflow Analyst,Tax Maintainer,Accounting Object Analyst,Tax Expert,AP Invoice Processor (GFT Job Role),Finance Integration Display and Reporting,MI Query Writer,Regional Maximo Labor Data Steward,Regional Backbone Administrator
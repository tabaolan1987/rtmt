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
contractor : Contractor? (Y/N)
#changeNetworkLevel : Change Network Level
#dofa : DofA (Y/N)
#dofaType : DofA Type  (Indent, Commitment, Sourcing)
#region : Region
#siteLocation : Maximo Site Location
#purchasingOrg : Purchasing Org
SFunction : Standard Function
SdSubFunction : Standard Sub Function
STeam : Standard Team
spare1 : Optional Field 1
spare2 : Optional Field 2
spare3 : Optional Field 3
spare4 : Optional Field 4
spare5 : Optional Field 5
spare6 : Optional Field 6
spare7 : Optional Field 7
spare8 : Optional Field 8
spare9 : Optional Field 9
spare10 : Optional Field 10
spare11 : Optional Field 11
spare12 : Optional Field 12
spare13 : Optional Field 13
spare14 : Optional Field 14
spare15 : Optional Field 15
spare16 : Optional Field 16
spare17 : Optional Field 17
spare18 : Optional Field 18
spare19 : Optional Field 19
spare20 : Optional Field 20
# Custom import format:   "insert query" | "check value" : "list of column to check"
# INSERT INTO user_data_mapping_role(idUserdata, idRegion, idBpRoleStandard) SELECT top 1 [ntid],idRegion,idBpRoleStandard FROM BpRoleStandard, Region where BpRoleStandardName = [value] And RegionName = [region_name] | y : SRM Lead Requester,Standard Desktop Confirmation Requester,Central Desktop Confirmation Requester,Backbone Reviewer/Approver,Procurement Display & Reporting,Contract Display & Reporting,Service Entry Sheet Creator,Operational Buyer,Contract Owner,Procurement Catalogue Processor,PSCM Specialist,PSCM Team Leader,Sourcing Coordinator,T&C Suppressor,Confidential Contracts Management,PSCM Approver,Bidder Creator,Goods Receiver,Title Transfer Receiver,Warehouse / Logistics Specialist,MRP Specialist,Strategic Materials Planner,Demand Planning MRP Specialist,VMI Administrator,Third Party Inventory Administrator,Inventory Scrapper,Inventory Transfer Specialist,Stock Count Administrator,Stock Count Variance Processor,Inventory Requestor,Materials Fabrication Requestor,MM Financial Approver,Materials Management Display & Reporting,CRP Processor,Cost Allocation Administrator (GoM),Shipment Specialist,Rental Specialist,Materials Expediter,Inventory Optimisation Analyst,POQR Library Administrator,POQR Document Reviewer,POQR Approver,WM Supervisor,WM Scheduler,WM Planner,WM Advanced Planner 1,WM Advanced Planner 2,WM Technician,WM Mobile Technician,WM Microsoft Project,WM Senior Leadership,WM Display & Reporting,WM Local Work Management Administrator,WM Regional Work Management Administrator,Master Data Administrator - Item,Master Data Administrator - Warehouse,Master Data Administrator - BOM/Product Structure,Master Data Administrator - Service,MDM - Global Data Steward,MDM - Local Data Steward,MDM - Display,Master Data Administrator - PSCM,Cost Approver Maintainer,Vendor Data Requestor (Egypt only),Vendor Maintainer - HSSE,GWO Data Maintainer,GWO Data Display,Order Settlement,Invoice Exception and Workflow Analyst,Tax Maintainer,Accounting Object Analyst,Tax Expert,AP Invoice Processor (GFT Job Role),Finance Integration Display and Reporting,MI Query Writer,Regional Maximo Labor Data Steward,Regional Backbone Administrator
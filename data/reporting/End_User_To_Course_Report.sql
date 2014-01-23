Select{
%
NTID,Full Name,First Name,Surname,OMS Function
|
select udata.ntid,(udata.lname + ',' + udata.fname) as [Fullname],
udata.fname,udata.lname,udata.omsSubfunction
from (((((((user_data as udata 
inner join specialism as sp)
on sp.SpecialismName = udata.specialism
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.idActivity)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.idActivity = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse) 
group by cou.courseId,udata.fname,udata.lname,udata.omsSubfunction,udata.ntid
%}
=====

=====
Select{%
Course Area,Course ID,Course Name,Course Type,Course Duration,Primary / Secondary,Pre / Post Live
|
select cou.courseArena,cou.courseId,
cou.courseTitle,cou.courseType,cou.courseDuration,
'' as [Primary / Secondary],cou.courseDelivery
from (((((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.idActivity)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.idActivity = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse )
where udata.ntid = UD.ntid
group by cou.courseId, cou.courseTitle, cou.courseType,
cou.courseDuration,cou.courseDelivery,cou.courseArena
%}
FROM user_data AS UD ORDER BY UD.ntid

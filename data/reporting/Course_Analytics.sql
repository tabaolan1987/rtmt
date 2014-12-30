Course Analytics
=======
SELECT Q_EDULBBJR2.courseId, Q_EDULBBJR2.courseTitle, Q_EDULBBJR2.courseType, 
Q_EDULBBJR2.courseDelivery, Q_EDULBBJR2.courseDuration, Q_EDULBBJR2.CountOfntid AS [Number of Users], 
[Q_EDULBBJR2]![courseDuration]*[Q_EDULBBJR2]![CountOfntid] AS [Total Hours]
FROM (SELECT Q_EDULBBJR1.BpRoleStandardName, course.courseId, course.courseTitle, 
course.courseType, course.courseDelivery, course.courseDuration, 
Count(Q_EDULBBJR1.ntid) as CountOfntid
FROM ((SELECT user_data.ntid, user_data.gpid, user_data.fname, user_data.SdSubFunction, 
user_data.country, user_data.lname, user_data_mapping_role.idRegion, 
Functions.id, user_data.mapped_bb_job_roles, user_data.blueprintRole, 
BpRoleStandard.BpRoleStandardName, BpRoleStandard.id
FROM ((user_data INNER JOIN user_data_mapping_role ON user_data.ntid = user_data_mapping_role.idUserdata) INNER JOIN BpRoleStandard ON user_data_mapping_role.idBpRoleStandard = BpRoleStandard.id) 
INNER JOIN Functions ON user_data.SFunction = Functions.nameFunction
WHERE (((user_data_mapping_role.idRegion)='(%RG_NAME%)') 
AND ((user_data.Deleted)=False) 
AND ((BpRoleStandard.Deleted)=False))
ORDER BY user_data.ntid) as Q_EDULBBJR1 INNER JOIN CourseMappingBpRoleStandard 
ON (Q_EDULBBJR1.Functions.id = CourseMappingBpRoleStandard.idFunction) 
AND (Q_EDULBBJR1.idRegion = CourseMappingBpRoleStandard.idRegion) 
AND (Q_EDULBBJR1.BpRoleStandard.id = CourseMappingBpRoleStandard.idBpRole)) 
INNER JOIN course ON (Q_EDULBBJR1.Functions.id = course.idFunction) 
AND (Q_EDULBBJR1.idRegion = course.idRegion) 
AND (CourseMappingBpRoleStandard.idCourse = course.id)
GROUP BY Q_EDULBBJR1.BpRoleStandardName, course.courseId, course.courseTitle, course.courseType, course.courseDelivery, 
course.courseDuration, Q_EDULBBJR1.idRegion, CourseMappingBpRoleStandard.Deleted, 
course.Deleted
HAVING (((CourseMappingBpRoleStandard.Deleted)=False) AND ((course.Deleted)=False))
ORDER BY course.courseId, Count(Q_EDULBBJR1.ntid)
) as Q_EDULBBJR2
GROUP BY Q_EDULBBJR2.courseId, Q_EDULBBJR2.courseTitle, Q_EDULBBJR2.courseType, 
Q_EDULBBJR2.courseDelivery, Q_EDULBBJR2.courseDuration, Q_EDULBBJR2.CountOfntid, 
[Q_EDULBBJR2]![courseDuration]*[Q_EDULBBJR2]![CountOfntid]
=========
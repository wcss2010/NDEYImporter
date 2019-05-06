using NDEYImporter.DB;
using Noear.Weed;
using System;
using System.Collections.Generic;
using System.Text;

namespace NDEYImporter.Util
{
    /// <summary>
    /// 数据库导入类，主要用于将NDEY数据库导入到汇总数据库
    /// </summary>
    public class DBImporter
    {
        /// <summary>
        /// 将NDEY数据添加到本地数据库中
        /// </summary>
        /// <param name="sourceFile">NDEY数据文本</param>        
        /// <returns>ProjectID</returns>        
        public static string appendNdeyToLocalDB(string sourceFile)
        {
            //ProjectID
            string projID = Guid.NewGuid().ToString();

            //SQLite数据库工厂
            System.Data.SQLite.SQLiteFactory factory = new System.Data.SQLite.SQLiteFactory();

            //NDEY数据库连接
            Noear.Weed.DbContext context = new Noear.Weed.DbContext("main", "Data Source = " + sourceFile, factory);

            try
            {
                //项目信息 （Old）BaseInfor --- （New）Project

                //旧的工程信息数据
                DataList dlOldProjects = context.table("BaseInfor").select("*").getDataList();

                //要添加到本地数据库的记录
                DataItem diNewProject = new DataItem();
                diNewProject.set("ID", projID);
                diNewProject.set("UserName", dlOldProjects.getRow(0).get("UserName"));
                diNewProject.set("Sex", dlOldProjects.getRow(0).get("Sex"));
                diNewProject.set("Birthdate", dlOldProjects.getRow(0).get("Birthdate"));
                diNewProject.set("Degree", dlOldProjects.getRow(0).get("Degree"));
                diNewProject.set("JobTitle", dlOldProjects.getRow(0).get("JobTitle"));
                diNewProject.set("UnitPosition", dlOldProjects.getRow(0).get("UnitPosition"));
                diNewProject.set("MainResearch", dlOldProjects.getRow(0).get("MainResearch"));
                diNewProject.set("ResultConfig", dlOldProjects.getRow(0).get("IDCardName"));
                diNewProject.set("CardNo", dlOldProjects.getRow(0).get("CardNo"));
                diNewProject.set("OfficePhones", dlOldProjects.getRow(0).get("OfficePhones"));
                diNewProject.set("MobilePhone", dlOldProjects.getRow(0).get("MobilePhone"));
                diNewProject.set("EMail", dlOldProjects.getRow(0).get("EMail"));
                diNewProject.set("UnitID", dlOldProjects.getRow(0).get("UnitID"));
                diNewProject.set("UnitName", dlOldProjects.getRow(0).get("UnitName"));
                diNewProject.set("UnitNormal", dlOldProjects.getRow(0).get("UnitNormal"));
                diNewProject.set("UnitIDCard", dlOldProjects.getRow(0).get("UnitIDCard"));
                diNewProject.set("UnitContacts", dlOldProjects.getRow(0).get("UnitContacts"));
                diNewProject.set("UnitAddress", dlOldProjects.getRow(0).get("UnitAddress"));
                diNewProject.set("UnitForORG", dlOldProjects.getRow(0).get("UnitForORG"));
                diNewProject.set("UnitProperties", dlOldProjects.getRow(0).get("UnitProperties"));
                diNewProject.set("UnitContactsPhone", dlOldProjects.getRow(0).get("UnitContactsPhone"));
                diNewProject.set("ProjectSecret", dlOldProjects.getRow(0).get("ProjectSecret"));
                diNewProject.set("ProjectName", dlOldProjects.getRow(0).get("ProjectName"));
                diNewProject.set("ApplicationType", dlOldProjects.getRow(0).get("ApplicationType"));
                diNewProject.set("ApplicationArea", dlOldProjects.getRow(0).get("ApplicationArea"));
                diNewProject.set("ProjectTD", dlOldProjects.getRow(0).get("ProjectTD"));
                diNewProject.set("ProjectMRD", dlOldProjects.getRow(0).get("ProjectMRD"));
                diNewProject.set("ProjectBaseT", dlOldProjects.getRow(0).get("ProjectBaseT"));
                diNewProject.set("ProjectKeyWord", dlOldProjects.getRow(0).get("ProjectKeyWord"));
                diNewProject.set("ProjectBrief", dlOldProjects.getRow(0).get("ProjectBrief"));
                diNewProject.set("ProjectLimitT", dlOldProjects.getRow(0).get("ProjectLimitT"));
                diNewProject.set("ProjectRFA", dlOldProjects.getRow(0).get("ProjectRFA"));
                diNewProject.set("ProjectRFA1", dlOldProjects.getRow(0).get("ProjectRFA1"));
                diNewProject.set("ProjectRFA1_1", dlOldProjects.getRow(0).get("ProjectRFA1_1"));
                diNewProject.set("ProjectRFA1_1_1", dlOldProjects.getRow(0).get("ProjectRFA1_1_1"));
                diNewProject.set("ProjectRFA1_1_2", dlOldProjects.getRow(0).get("ProjectRFA1_1_2"));
                diNewProject.set("ProjectRFA1_1_3", dlOldProjects.getRow(0).get("ProjectRFA1_1_3"));
                diNewProject.set("ProjectRFA1_2", dlOldProjects.getRow(0).get("ProjectRFA1_2"));
                diNewProject.set("ProjectRFA1_3", dlOldProjects.getRow(0).get("ProjectRFA1_3"));
                diNewProject.set("ProjectRFA1_4", dlOldProjects.getRow(0).get("ProjectRFA1_4"));
                diNewProject.set("ProjectRFA1_5", dlOldProjects.getRow(0).get("ProjectRFA1_5"));
                diNewProject.set("ProjectRFA1_6", dlOldProjects.getRow(0).get("ProjectRFA1_6"));
                diNewProject.set("ProjectRFA1_7", dlOldProjects.getRow(0).get("ProjectRFA1_7"));
                diNewProject.set("ProjectRFA1_8", dlOldProjects.getRow(0).get("ProjectRFA1_8"));
                diNewProject.set("ProjectRFA1_9", dlOldProjects.getRow(0).get("ProjectRFA1_9"));
                diNewProject.set("ProjectRFA2", dlOldProjects.getRow(0).get("ProjectRFA2"));
                diNewProject.set("ProjectRFA2_1", dlOldProjects.getRow(0).get("ProjectRFA2_1"));
                diNewProject.set("UnitRecommend", dlOldProjects.getRow(0).get("UnitRecommend"));
                diNewProject.set("UnitRecommendOName", dlOldProjects.getRow(0).get("UnitRecommendOName"));
                diNewProject.set("ExpertName1", dlOldProjects.getRow(0).get("ExpertName1"));
                diNewProject.set("ExpertArea1", dlOldProjects.getRow(0).get("ExpertArea1"));
                diNewProject.set("ExpertUnitPosition1", dlOldProjects.getRow(0).get("ExpertUnitPosition1"));
                diNewProject.set("ExpertUnit1", dlOldProjects.getRow(0).get("ExpertUnit1"));
                diNewProject.set("ExpertRecommend1", dlOldProjects.getRow(0).get("ExpertRecommend1"));
                diNewProject.set("ExpertRecommend1OName", dlOldProjects.getRow(0).get("ExpertRecommend1OName"));
                diNewProject.set("ExpertName2", dlOldProjects.getRow(0).get("ExpertName2"));
                diNewProject.set("ExpertArea2", dlOldProjects.getRow(0).get("ExpertArea2"));
                diNewProject.set("ExpertUnitPosition2", dlOldProjects.getRow(0).get("ExpertUnitPosition2"));
                diNewProject.set("ExpertUnit2", dlOldProjects.getRow(0).get("ExpertUnit2"));
                diNewProject.set("ExpertRecommend2", dlOldProjects.getRow(0).get("ExpertRecommend2"));
                diNewProject.set("ExpertRecommend2OName", dlOldProjects.getRow(0).get("ExpertRecommend2OName"));
                diNewProject.set("ExpertName3", dlOldProjects.getRow(0).get("ExpertName3"));
                diNewProject.set("ExpertArea3", dlOldProjects.getRow(0).get("ExpertArea3"));
                diNewProject.set("ExpertUnitPosition3", dlOldProjects.getRow(0).get("ExpertUnitPosition3"));
                diNewProject.set("ExpertUnit3", dlOldProjects.getRow(0).get("ExpertUnit3"));
                diNewProject.set("ExpertRecommend3", dlOldProjects.getRow(0).get("ExpertRecommend3"));
                diNewProject.set("ExpertRecommend3OName", dlOldProjects.getRow(0).get("ExpertRecommend3OName"));
                ConnectionManager.Context.table("Project").insert(diNewProject);

                //单位信息 （Old）UnitInfor --- （New）Unit

                //学术经历 （Old）AcademicPost --- （New）ResearchExperience

                //工作经历 （Old）WorkExperience --- （New）WorkExperience

                //教育经历 （Old）Education --- （New）EducationExperience

                //获得科技奖项经历 （Old）TechnologyAwards --- （New）TechnologyAwardsExperience

                //入选人才计划经历 （Old）TalentsPlan --- （New）TalentsPlanExperience

                //代表性论著经历 （Old）RTreatises --- （New）TreatisesExperience

                //承担国防项目经历 （Old）NDProject --- （New）NationalDefenseProjectExperience

                //获得国防专利经历 （Old）NDPatent --- （New）NationalDefensePatentExperience 
            }
            finally
            {
                factory.Dispose();
                context = null;
            }

            return projID;
        }
    }
}
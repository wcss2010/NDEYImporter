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
        public static void appendNdeyToLocalDB(string sourceFile)
        {
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
                diNewProject.set("ID", Guid.NewGuid().ToString());
                diNewProject.set("UserName", dlOldProjects.getRow(0).get(""));
                diNewProject.set("Sex", dlOldProjects.getRow(0).get(""));
                diNewProject.set("Birthdate", dlOldProjects.getRow(0).get(""));
                diNewProject.set("Degree", dlOldProjects.getRow(0).get(""));
                diNewProject.set("JobTitle", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitPosition", dlOldProjects.getRow(0).get(""));
                diNewProject.set("MainResearch", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ResultConfig", dlOldProjects.getRow(0).get(""));
                diNewProject.set("CardNo", dlOldProjects.getRow(0).get(""));
                diNewProject.set("OfficePhones", dlOldProjects.getRow(0).get(""));
                diNewProject.set("MobilePhone", dlOldProjects.getRow(0).get(""));
                diNewProject.set("EMail", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitID", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitName", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitNormal", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitIDCard", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitContacts", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitAddress", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitForORG", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitProperties", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitContactsPhone", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectSecret", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectName", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ApplicationType", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ApplicationArea", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectTD", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectMRD", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectBaseT", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectKeyWord", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectBrief", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectLimitT", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_1_1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_1_2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_1_3", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_3", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_4", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_5", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_6", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_7", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_8", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA1_9", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ProjectRFA2_1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitRecommend", dlOldProjects.getRow(0).get(""));
                diNewProject.set("UnitRecommendOName", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertName1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertArea1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertUnitPosition1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertUnit1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertRecommend1", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertRecommend1OName", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertName2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertArea2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertUnitPosition2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertUnit2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertRecommend2", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertRecommend2OName", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertName3", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertArea3", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertUnitPosition3", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertUnit3", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertRecommend3", dlOldProjects.getRow(0).get(""));
                diNewProject.set("ExpertRecommend3OName", dlOldProjects.getRow(0).get(""));

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
        }
    }
}
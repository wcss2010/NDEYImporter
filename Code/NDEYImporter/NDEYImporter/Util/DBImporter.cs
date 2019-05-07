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
        /// 导入数据库
        /// </summary>
        /// <param name="sourceFile">NDEY数据文件</param>
        /// <returns>ProjectID</returns>
        public static string importDB(string sourceFile)
        {
            //项目ID
            string projectID = string.Empty;

            //SQLite数据库工厂
            System.Data.SQLite.SQLiteFactory factory = new System.Data.SQLite.SQLiteFactory();

            //NDEY数据库连接
            Noear.Weed.DbContext context = new Noear.Weed.DbContext("main", "Data Source = " + sourceFile, factory);

            try
            {
                //项目信息 （Old）BaseInfor --- （New）Project

                //旧的工程信息数据
                DataList dlOldProjects = context.table("BaseInfor").select("*").getDataList();

                //判断数据库是否有内容
                if (dlOldProjects.getRowCount() >= 1)
                {
                    //项目名称
                    string projectName = dlOldProjects.getRow(0).getString("ProjectName");
                    //申请人
                    string projectCreater = dlOldProjects.getRow(0).getString("UserName");

                    //查找项目索引
                    DataList dlCatalogs = ConnectionManager.Context.table("Catalog").where("ProjectName='" + projectName + "' and ProjectCreater='" + projectCreater + "'").select("*").getDataList();

                    //判断是否需要新建索引
                    if (dlCatalogs.getRowCount() >= 1)
                    {
                        //不需要新建

                        //当前项目ID
                        projectID = dlCatalogs.getRow(0).getString("ProjectID");

                        //清理项目数据
                        clearProjectData(projectID);

                        //插入数据到本地库中
                        projectID = insertToLocalDB(projectID, sourceFile);

                        //更新索引中的项目ID
                        ConnectionManager.Context.table("Catalog").set("ProjectID", projectID).where("ProjectName='" + projectName + "' and ProjectCreater='" + projectCreater + "'").update();
                    }
                    else
                    {
                        //需要新建
                        
                        //插入数据到本地库中
                        //新项目ID
                        projectID = insertToLocalDB(Guid.NewGuid().ToString(), sourceFile);

                        //项目单位ID
                        string projectCreaterUnitID = ConnectionManager.Context.table("Project").where("ID='" + projectID + "'").select("UnitID").getValue<string>(string.Empty);

                        //创建索引
                        DataItem diCatalog = new DataItem();
                        diCatalog.set("ProjectID", projectID);
                        diCatalog.set("ProjectNumber", getProjectNumbers());
                        diCatalog.set("ProjectName", projectName);
                        diCatalog.set("ProjectCreater", projectCreater);
                        diCatalog.set("ProjectCreaterUnitID", projectCreaterUnitID);
                        ConnectionManager.Context.table("Catalog").insert(diCatalog);
                    }
                }
                else
                {
                    //空数据库
                    throw new Exception("对不起，数据库为空！");
                }
            }
            finally
            {
                factory.Dispose();
                context = null;
            }

            return projectID;
        }

        /// <summary>
        /// 生成项目编号
        /// </summary>
        /// <returns></returns>
        private static string getProjectNumbers()
        {
            return string.Empty;
        }

        /// <summary>
        /// 清理本地数据库中的指定项目ID下的所有数据
        /// </summary>
        /// <param name="projID">项目ID</param>
        private static void clearProjectData(string projID)
        {
            //清理项目信息
            ConnectionManager.Context.table("Project").where("ID='" + projID + "'").delete();

            //清理学术经历
            ConnectionManager.Context.table("ResearchExperience").where("ProjectID='" + projID + "'").delete();
            //清理工作经历
            ConnectionManager.Context.table("WorkExperience").where("ProjectID='" + projID + "'").delete();
            //清理教育经历
            ConnectionManager.Context.table("EducationExperience").where("ProjectID='" + projID + "'").delete();
            //清理获得科技奖项经历
            ConnectionManager.Context.table("TechnologyAwardsExperience").where("ProjectID='" + projID + "'").delete();
            //清理入选人才计划经历
            ConnectionManager.Context.table("TalentsPlanExperience").where("ProjectID='" + projID + "'").delete();
            //清理代表性论著经历
            ConnectionManager.Context.table("TreatisesExperience").where("ProjectID='" + projID + "'").delete();
            //清理承担国防项目经历
            ConnectionManager.Context.table("NationalDefenseProjectExperience").where("ProjectID='" + projID + "'").delete();
            //清理获得国防专利经历
            ConnectionManager.Context.table("NationalDefensePatentExperience ").where("ProjectID='" + projID + "'").delete();
                        
            //单位ID
            string unitId = ConnectionManager.Context.table("Catalog").where("ProjectID='" + projID + "'").select("ProjectCreaterUnitID").getValue<string>(string.Empty);
            //使用次数
            long unitUseCount = ConnectionManager.Context.table("Catalog").where("ProjectCreaterUnitID='" + unitId + "'").select("count(*)").getValue<long>(0);
            //判断是否可以删除
            if (unitUseCount <= 1)
            {
                //只有当这个单位只有一次引用时才能删除
                //清理单位信息
                ConnectionManager.Context.table("Unit").where("ID='" + unitId + "' and IsUserAdded = 1").delete();
            }
        }

        /// <summary>
        /// 将NDEY数据添加到本地数据库中
        /// </summary>
        /// <param name="projID">项目ID</param>
        /// <param name="sourceFile">NDEY数据文件</param>        
        /// <returns>ProjectID</returns>        
        private static string insertToLocalDB(string projID, string sourceFile)
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

                //判断数据库是否有内容
                if (dlOldProjects.getRowCount() >= 1)
                {
                    //有效数据库

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
                    //这里的逻辑是先看看当前用的单位是不是自定义的，如果不是，就无需添加，如果是则检查之前有没有添加过相同的记录，如果添加过，就替换掉这条记录，如果没有直接添加
                    //旧的单位信息数据
                    DataList dlUnitList = context.table("UnitInfor").where("ID = '" + diNewProject.getString("UnitID") + "'").select("*").getDataList();

                    //判断是否选择了单位信息
                    if (dlUnitList.getRowCount() >= 1)
                    {
                        //有选择单位信息
                        
                        //判断是否为自定义单位
                        if (dlUnitList.getRow(0).getInt("IsUserAdded") == 1)
                        {
                            //自定义单位

                            //检查本地库中有没有这条记录
                            DataList dlLocalCustomUnit = ConnectionManager.Context.table("Unit").where("UnitBankNo='" + dlUnitList.getRow(0).getString("UnitBankNo") + "'").select("*").getDataList();

                            //添加或更新的单位信息
                            DataItem diNewUnit = new DataItem();
                            diNewUnit.set("ID", Guid.NewGuid().ToString());
                            diNewUnit.set("UnitName", dlUnitList.getRow(0).get("UnitName"));
                            diNewUnit.set("UnitType", dlUnitList.getRow(0).get("UnitType"));
                            diNewUnit.set("UnitBankUser", dlUnitList.getRow(0).get("UnitBankUser"));
                            diNewUnit.set("UnitBankName", dlUnitList.getRow(0).get("UnitBankName"));
                            diNewUnit.set("UnitBankNo", dlUnitList.getRow(0).get("UnitBankNo"));
                            diNewUnit.set("IsUserAdded", dlUnitList.getRow(0).get("IsUserAdded"));

                            //判断是不是添加过这条自定义单位
                            if (dlLocalCustomUnit.getRowCount() >= 1)
                            {
                                //删除掉旧的记录
                                ConnectionManager.Context.table("Unit").where("ID='" + dlLocalCustomUnit.getRow(0).getString("ID") + "'").delete();

                                //更新Project的UnitID
                                ConnectionManager.Context.table("Project").set("UnitID", diNewUnit.getString("ID")).where("UnitID = '" + dlLocalCustomUnit.getRow(0).getString("ID") + "'").update();

                                //更新Catalog的ProjectCreaterUnitID
                                ConnectionManager.Context.table("Catalog").set("ProjectCreaterUnitID", diNewUnit.getString("ID")).where("ProjectCreaterUnitID = '" + dlLocalCustomUnit.getRow(0).getString("ID") + "'").update();
                            }

                            //添加单位数据
                            ConnectionManager.Context.table("Unit").insert(diNewUnit);

                            //更新本项目的Project的UnitID
                            ConnectionManager.Context.table("Project").set("UnitID", diNewUnit.getString("ID")).where("ID = '" + projID + "'").update();

                            //更新本项目的Catalog的ProjectCreaterUnitID
                            ConnectionManager.Context.table("Catalog").set("ProjectCreaterUnitID", diNewUnit.getString("ID")).where("ProjectID = '" + projID + "'").update();
                        }
                        else
                        {
                            //更新本项目的Catalog的ProjectCreaterUnitID
                            ConnectionManager.Context.table("Catalog").set("ProjectCreaterUnitID", dlUnitList.getRow(0).getString("ID")).where("ProjectID = '" + projID + "'").update();
                        }
                    }
                    else
                    {
                        //没有选择单位信息
                        throw new Exception("对不起，没有选择单位信息！");
                    }

                    //学术经历 （Old）AcademicPost --- （New）ResearchExperience

                    //旧的学术经历数据
                    DataList dlOldResearchList = context.table("AcademicPost").select("*").getDataList();

                    //循环插入学术经历数据到本地库
                    foreach (DataItem di in dlOldResearchList.getRows())
                    {
                        //添加学术经历数据
                        DataItem diNewResearch = new DataItem();
                        diNewResearch.set("ProjectID", projID);
                        diNewResearch.set("ID", Guid.NewGuid().ToString());
                        diNewResearch.set("StartDate", di.get("AcademicPostSDate"));
                        diNewResearch.set("EndDate", di.get("AcademicPostEDate"));
                        diNewResearch.set("Org", di.get("AcademicPostOrg"));
                        diNewResearch.set("Job", di.get("AcademicPostContent"));
                        diNewResearch.set("RecordOrder", di.get("AcademicPostOrder"));
                        ConnectionManager.Context.table("ResearchExperience").insert(diNewResearch);
                    }

                    //工作经历 （Old）WorkExperience --- （New）WorkExperience

                    //旧的工作经历数据
                    DataList dlOldWorkList = context.table("WorkExperience").select("*").getDataList();

                    //循环插入工作经历数据到本地库
                    foreach (DataItem di in dlOldWorkList.getRows())
                    {
                        //添加工作经历数据
                        DataItem diNewWork = new DataItem();
                        diNewWork.set("ProjectID", projID);
                        diNewWork.set("ID", Guid.NewGuid().ToString());
                        diNewWork.set("StartDate", di.get("WorkExperienceSDate"));
                        diNewWork.set("EndDate", di.get("WorkExperienceEDate"));
                        diNewWork.set("Org", di.get("WorkExperienceOrg"));
                        diNewWork.set("Job", di.get("WorkExperienceContent"));
                        diNewWork.set("RecordOrder", di.get("WorkExperienceOrder"));
                        ConnectionManager.Context.table("WorkExperience").insert(diNewWork);
                    }


                    //教育经历 （Old）Education --- （New）EducationExperience

                    //旧的教育经历数据
                    DataList dlOldEducationList = context.table("Education").select("*").getDataList();

                    //循环插入教育经历数据到本地库
                    foreach (DataItem di in dlOldEducationList.getRows())
                    {
                        //添加教育经历数据
                        DataItem diNewEducation = new DataItem();
                        diNewEducation.set("ProjectID", projID);
                        diNewEducation.set("ID", Guid.NewGuid().ToString());
                        diNewEducation.set("StartDate", di.get("EducationSDate"));
                        diNewEducation.set("EndDate", di.get("EducationEDate"));
                        diNewEducation.set("Org", di.get("EducationOrg"));
                        diNewEducation.set("Major", di.get("EducationMajor"));
                        diNewEducation.set("Degree", di.get("EducationDegree"));
                        diNewEducation.set("RecordOrder", di.get("EducationOrder"));
                        ConnectionManager.Context.table("EducationExperience").insert(diNewEducation);
                    }

                    //获得科技奖项经历 （Old）TechnologyAwards --- （New）TechnologyAwardsExperience

                    //旧的获得科技奖项经历数据
                    DataList dlOldTechnologyAwardsList = context.table("TechnologyAwards").select("*").getDataList();

                    //循环插入获得科技奖项经历数据到本地库
                    foreach (DataItem di in dlOldTechnologyAwardsList.getRows())
                    {
                        //添加获得科技奖项经历数据
                        DataItem diNewTechnologyAwards = new DataItem();
                        diNewTechnologyAwards.set("ProjectID", projID);
                        diNewTechnologyAwards.set("ID", Guid.NewGuid().ToString());
                        diNewTechnologyAwards.set("Name", di.get("TechnologyAwardsPName"));
                        diNewTechnologyAwards.set("Date", di.get("TechnologyAwardsYear"));
                        diNewTechnologyAwards.set("TypeLevel", di.get("TechnologyAwardsTypeLevel"));
                        diNewTechnologyAwards.set("PDF", di.get("TechnologyAwardsPDF"));
                        diNewTechnologyAwards.set("PDFOName", di.get("TechnologyAwardsPDFOName"));
                        diNewTechnologyAwards.set("Ranking", di.get("TechnologyAwardsee"));
                        diNewTechnologyAwards.set("RecordOrder", di.get("TechnologyAwardsOrder"));
                        ConnectionManager.Context.table("TechnologyAwardsExperience").insert(diNewTechnologyAwards);
                    }

                    //入选人才计划经历 （Old）TalentsPlan --- （New）TalentsPlanExperience

                    //旧的入选人才计划经历数据
                    DataList dlOldTalentsPlanList = context.table("TalentsPlan").select("*").getDataList();

                    //循环插入入选人才计划经历数据到本地库
                    foreach (DataItem di in dlOldTalentsPlanList.getRows())
                    {
                        //添加入选人才计划经历数据
                        DataItem diNewTalentsPlan = new DataItem();
                        diNewTalentsPlan.set("ProjectID", projID);
                        diNewTalentsPlan.set("ID", Guid.NewGuid().ToString());
                        diNewTalentsPlan.set("Date", di.get("TalentsPlanDate"));
                        diNewTalentsPlan.set("Name", di.get("TalentsPlanName"));
                        diNewTalentsPlan.set("RA", di.get("TalentsPlanRA"));
                        diNewTalentsPlan.set("Outlay", di.get("TalentsPlanOutlay"));
                        diNewTalentsPlan.set("RecordOrder", di.get("TalentsPlanOrder"));
                        ConnectionManager.Context.table("TalentsPlanExperience").insert(diNewTalentsPlan);
                    }

                    //代表性论著经历 （Old）RTreatises --- （New）TreatisesExperience

                    //旧的代表性论著经历数据
                    DataList dlOldRTreatisesList = context.table("RTreatises").select("*").getDataList();

                    //循环插入代表性论著经历数据到本地库
                    foreach (DataItem di in dlOldRTreatisesList.getRows())
                    {
                        //添加代表性论著经历数据
                        DataItem diNewRTreatises = new DataItem();
                        diNewRTreatises.set("ProjectID", projID);
                        diNewRTreatises.set("ID", Guid.NewGuid().ToString());
                        diNewRTreatises.set("Type", di.get("RTreatisesType"));
                        diNewRTreatises.set("Name", di.get("RTreatisesName"));
                        diNewRTreatises.set("ORG", di.get("RTreatisesJournalTitle"));
                        diNewRTreatises.set("Date", di.get("RTreatisesRell"));
                        diNewRTreatises.set("Collection", di.get("RTreatisesCollection"));
                        diNewRTreatises.set("PDF", di.get("RTreatisesPDF"));
                        diNewRTreatises.set("PDFOName", di.get("RTreatisesPDFOName"));
                        diNewRTreatises.set("Ranking", di.get("RTreatisesAuthor"));
                        diNewRTreatises.set("RecordOrder", di.get("RTreatisesOrder"));
                        ConnectionManager.Context.table("TreatisesExperience").insert(diNewRTreatises);
                    }

                    //承担国防项目经历 （Old）NDProject --- （New）NationalDefenseProjectExperience

                    //旧的承担国防项目经历数据
                    DataList dlOldNDProjectList = context.table("NDProject").select("*").getDataList();

                    //循环插入承担国防项目经历数据到本地库
                    foreach (DataItem di in dlOldNDProjectList.getRows())
                    {
                        //添加承担国防项目经历数据
                        DataItem diNewNDProject = new DataItem();
                        diNewNDProject.set("ProjectID", projID);
                        diNewNDProject.set("ID", Guid.NewGuid().ToString());
                        diNewNDProject.set("StartDate", di.get("NDProjectSYear"));
                        diNewNDProject.set("EndDate", di.get("NDProjectEYear"));
                        diNewNDProject.set("Name", di.get("NDProjectName"));
                        diNewNDProject.set("Source", di.get("NDProjectSource"));
                        diNewNDProject.set("Outlay", di.get("NDProjectOutlay"));
                        diNewNDProject.set("TaskBySelf", di.get("NDProjectTaskBySelf"));
                        diNewNDProject.set("Ranking", di.get("NDProjectUserOrder"));
                        diNewNDProject.set("RecordOrder", di.get("NDProjectOrder"));
                        ConnectionManager.Context.table("NationalDefenseProjectExperience").insert(diNewNDProject);
                    }

                    //获得国防专利经历 （Old）NDPatent --- （New）NationalDefensePatentExperience 

                    //旧的获得国防专利经历数据
                    DataList dlOldNDPatentList = context.table("NDPatent").select("*").getDataList();

                    //循环插入获得国防专利经历数据到本地库
                    foreach (DataItem di in dlOldNDPatentList.getRows())
                    {
                        //添加获得国防专利经历数据
                        DataItem diNewNDPatent = new DataItem();
                        diNewNDPatent.set("ProjectID", projID);
                        diNewNDPatent.set("ID", Guid.NewGuid().ToString());
                        diNewNDPatent.set("PatentName", di.get("NDPatentName"));
                        diNewNDPatent.set("PatentType", di.get("NDPatentType"));
                        diNewNDPatent.set("PatentApprovalYear", di.get("NDPatentApprovalYear"));
                        diNewNDPatent.set("PatentNumber", di.get("NDPatentNumber"));
                        diNewNDPatent.set("PDF", di.get("NDPatentPDF"));
                        diNewNDPatent.set("PDFOName", di.get("NDPatentPDFOName"));
                        diNewNDPatent.set("Ranking", di.get("NDPatentApplicants"));
                        diNewNDPatent.set("RecordOrder", di.get("NDPatentOrder"));
                        ConnectionManager.Context.table("NationalDefensePatentExperience").insert(diNewNDPatent);
                    }
                }
                else
                {
                    //空数据库
                    throw new Exception("对不起，数据库为空！");
                }
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
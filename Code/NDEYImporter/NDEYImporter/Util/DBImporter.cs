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
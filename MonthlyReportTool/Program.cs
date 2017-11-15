﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //API.TFS.Utils.RetrieveTeamMemberList("");
            var prjlist = API.TFS.TeamProject.Project.RetrieveProjectList();
            foreach (var prj in prjlist)
            {
                Console.WriteLine(prj.Name);
                
                if (prj.Name.ToLower() == "bugs") continue;
                if (prj.Name.ToLower() != "fcp") continue;
                //var teamlist = API.TFS.TeamProject.Team.RetrieveTeamList(prj.Name);
                //foreach (var team in teamlist)
                //{
                //    Console.WriteLine("\t"+team.Name);
                //    var memlist = API.TFS.TeamProject.Member.RetrieveMemberList(prj.Name, team.Name);
                //    foreach (var mem in memlist)
                //    {
                //        Console.WriteLine("\t\t" + mem.DisplayName);
                //    }
                //}
                //continue;
                API.Office.Excel.Utility.BuildIterationReports(prj);


                Console.WriteLine("======================");
            }
            //API.Office.Excel.Utility.BuildIterationReports();

            //var features = API.TFS.Utils.GetAllFeaturesByIterations("TTP\\FYQ4\\Sprint35");

        }
    }
}

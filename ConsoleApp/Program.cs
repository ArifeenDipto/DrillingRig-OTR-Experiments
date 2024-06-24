using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DrillSIM_API.API;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            DSAPI Project = new DSAPI();

            Project.Simulate();

            while (Project.IsRunning)
            {
                // Run info data
            }

            Project.Exit();
        }
    }
}

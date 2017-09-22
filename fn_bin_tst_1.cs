using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace VMS.TPS
{
    public class Script
    {
        // This is the cunstructor, has nothing but produce an instance!! 
        public Script()
        {
        }

        private SelectBox selectDiag;
        private String selected;
        

        public void Execute(ScriptContext context  /*, System.Windows.Window window */)
        {
            List<String> my_list = new List<string>();

            Patient my_patient = context.Patient;
            foreach (Course course in my_patient.Courses)
            { my_list.Add(course.Id);}
            selectDiag = new SelectBox(my_list, "Course Id");
            selected = selectDiag.Get_Item();

            Course my_course = my_patient.Courses.Where(c => c.Id.Equals(selected)).First();
            my_list.Clear();
            foreach (PlanSetup plan in my_course.PlanSetups)
            { my_list.Add(plan.Id); }
            selectDiag = new SelectBox(my_list, "Plan Id");
            selected = selectDiag.Get_Item();
            
            PlanSetup my_plan= my_course.PlanSetups.Where(c => c.Id.Equals(selected)).First();

            System.Windows.MessageBox.Show(my_plan.Id);

        }

        



    }

}



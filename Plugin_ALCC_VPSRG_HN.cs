using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
// using Plugin_ALCC_VPSRG_HN;



namespace VMS.TPS
{
    
    class Script
    {
        /// <summary>
        /// This is the cunstructor, do nothing other than allowing instance creation
        /// </summary>
        public Script()
        {
        }

        /// <summary>
        /// Plugin_ALCC_VPSRG_HN runs as binary script on Eclipse. It requires a patient open on the context.
        /// Promts for course and plan selection (prepared for VicRaplidPlan H&&N) and produces an csv txt file
        /// with demografics and quality metrics of plan for ViC Rapid plan project.
        /// </summary>
        /// <param name="context"></param>
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnviroment enviroment*/)
        {
            // Variables to be used on the selection forms
            SelectBox selectDiag;
            String selected;
            Structure selected_strct;
            List<String> my_list = new List<string>();
            List<Tuple<String,String>> selected_structs = new List<Tuple<String, String>>();
            SelectOneStruct selectOneStruct;
            IEnumerable<Structure> set_of_structs;
            IEnumerable<Structure> partial_set_of_structs;
            String title;
            Patient my_patient = context.Patient;

            // Output string
            // this string will get the text reporting for import on
            // VPSRG Head and Neck case-tracking sheet V2.xlsm. 
            // First three colums and fifth cannot be populated
            String VPSRG_HN_track = ",,," + my_patient.Id + ",,";  
                                                                   

            // "Metric" is a constant with a unit, even a "relative" unit (%)
            // "Goal" is a "Metric" that acts as an objective 
            // Metrics
            DoseValue Dose_Metric = new DoseValue(0, "Gy"); // absolute dose in Gy
            DoseValue Rel_Dose_Metric = new DoseValue(0.0, "%"); // relative dose in %
            double Vol_Metric = new double(); // absolute volume in cm3
            double Rel_Vol_Metric = new double(); // relative volume in %
            // Goals
            DoseValue Dose_Goal = new DoseValue(0, "Gy"); // absolute dose in Gy
            DoseValue Rel_Dose_Goal = new DoseValue(0, "%"); // relative dose goal in %
           // double Vol_Goal = new double(); // absolute volume in cm3
           // double Rel_Vol_Goal = new double(); // relative volume goal in %
           // Misc.
           // bool are_there = new bool(); // for testing existency of objects
           // String text = null; // a needed text container
            String part_name = null; // for defining the string to search an structure
                                     // VolumePresentation Vol_present = VolumePresentation.AbsoluteCm3;
            DoseValue Abs_Dose = new DoseValue(0.0, "Gy");
            DoseValue Rel_Dose = new DoseValue(0.0, "%");


            //*************** Select course
            foreach (Course course in my_patient.Courses)
            { my_list.Add(course.Id); }
            selectDiag = new SelectBox(my_list, "Course Id");
            selected = selectDiag.Get_Item();
            Course my_course = my_patient.Courses.Where(c => c.Id.Equals(selected)).First();
            // Write output string: Col F(6th) - Specific H&N anatomy gets course Id as normal practice at ALCC
            // following 4 cols (G-J) cannot be populated
            VPSRG_HN_track = VPSRG_HN_track + my_course.Id + ", ,,,,";

            //*************** Select plan
            my_list.Clear();
            foreach (PlanSetup plan in my_course.PlanSetups)
            { my_list.Add(plan.Id); }
            selectDiag = new SelectBox(my_list, "Plan Id");
            selected = selectDiag.Get_Item();

            PlanSetup my_plan = my_course.PlanSetups.Where(c => c.Id.Equals(selected)).First();
            //Checked if treated: Col K(tret' start day) cannot be populated but indication is writen if treated or not
            if (my_plan.IsTreated)
            { VPSRG_HN_track = VPSRG_HN_track + "Treated,"; }
            else
            { VPSRG_HN_track = VPSRG_HN_track + "Not Treated,"; }

            // DATA
            // Col L: # dose levels (equals # of PTV structures)
            // Getting all the non-empty structures that contains "ptv" and not contains "ip"
            part_name = "ptv"; // lower case part of the name to look for
            IEnumerable<Structure> set_of_ptvs = my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains(part_name) & !s.Id.ToLower().Contains("ip") & !s.IsEmpty);
            int num_dose_levels = set_of_ptvs.Count();
            // wite Col L data
            VPSRG_HN_track = VPSRG_HN_track + num_dose_levels.ToString() + ",";

            // Col M: Prescr. dose (in the units used in plan)
            DoseValue Dose_prescr = my_plan.TotalPrescribedDose;
            VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_prescr.Dose,2).ToString() + ",";

            // Col N: Percent isodose prescribed
            double Perc_Dose_prescr = my_plan.PrescribedPercentage * 100.0; // to have this in % with % unit
            VPSRG_HN_track = VPSRG_HN_track + Math.Round(Perc_Dose_prescr,1).ToString() + ",";  // Not use % for excell

            // Col O: Total MU
            double mu = 0.0;
            foreach (Beam beam in my_plan.Beams)
            { // folowing if for not getting NaN from setup beams
                if (!Double.IsNaN(beam.Meterset.Value))
                {
                    mu = mu + beam.Meterset.Value;
                }
            }
            VPSRG_HN_track = VPSRG_HN_track + Math.Round(mu,1).ToString() + ",";

            // Col P: IMRT or VMAT, following Col Q cannot be populated, skipping Col R
            String vmat = "VMAT";
            foreach (Beam b in my_plan.Beams)
            {
                if (!Double.IsNaN(b.Meterset.Value) && !(b.MLCPlanType.ToString() == "VMAT"))
                {
                    vmat = "IMRT";
                }
            }
            VPSRG_HN_track = VPSRG_HN_track + vmat + ", , ,"; // spkipping Col Q - Col R

            // GOALS evaluation

            // Col S: Brainstem *************************************************************************************************
            // Getting all the structures containing "stem" 
            part_name = "stem"; // lower case part of the name to look for
            title = "Brain Stem:";
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) & !s.IsEmpty);
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count()>1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose,2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title+" - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Col T: Cord ******************************************************************************************************
            // Getting all the structures containing "cord" and NOT "prv" nor "ip"
            part_name = "cord"; // lower case part of the name to look for
            title = "Spinal Cord:";
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) & !s.IsEmpty &&
                    !s.Id.ToLower().Contains("prv") && !s.Id.ToLower().Contains("ip"));
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Col U: Cord PRV **************************************************************************************************
            // Getting all the structures containing "cord" AND "prv" or "ip"
            part_name = "cord"; // lower case part of the name to look for
            title = "Spinal Cord PRV:";
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) &&
                    (s.Id.ToLower().Contains("PRV") || s.Id.ToLower().Contains("ip")) && !s.IsEmpty);
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            
            // Col V: Larynx ***************************************************************************************************
            // Getting all the structures containing "larynx"
            part_name = "larynx"; // lower case part of the name to look for
            title = "Larynx:";
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) & !s.IsEmpty);
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Inner ear Lt and Rt **********************************************************************************************
            // Getting all the structures containing "ear"
            part_name = "ear"; // lower case part of the name to look for
            partial_set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name));

            // COl W: Inner ear Lt
            set_of_structs = partial_set_of_structs.Where(s => s.Id.ToLower().Contains("l") & !s.IsEmpty);
            title = "Inner ear Lt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // COl X: Inner ear Rt
            set_of_structs = partial_set_of_structs.Where(s => !s.Id.ToLower().Contains("l") & !s.IsEmpty);
            title = "Inner ear Rt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }
            
            // Lens Lt and Rt **********************************************************************************************
            // Getting all the structures containing "lens" 
            part_name = "lens"; // lower case part of the name to look for
            partial_set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name));

            // COl Y: Lens Lt
            set_of_structs = partial_set_of_structs.Where(s => !s.Id.ToLower().Contains("r") & !s.IsEmpty);
            title = "Lens Lt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // COl Z: Lens Rt
            set_of_structs = partial_set_of_structs.Where(s => s.Id.ToLower().Contains("r") & !s.IsEmpty);
            title = "Lens Rt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }
            
            // Mandible ***************************************************************************************************
            // Getting all the structures containing "mandible" 
            part_name = "mandible"; // lower case part of the name to look for
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) & !s.IsEmpty);
            title = "Mandible:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // Col AA: D_Max, Col AB: V{TotalDose} [%]
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
                // Col AA D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                // Col AB V{TotalDose} [%]
                Rel_Vol_Metric = ALCC_QM.V_X_report(my_plan, selected_strct, Dose_prescr, VolumePresentation.Relative);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Rel_Vol_Metric, 2).ToString() + ",";// Not use % symbol for Excel
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN, Nan,"; // Cols AA, AB
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Col AC: Optic Chiasm **********************************************************************************************
            // Getting all the structures containing "chiasm" 
            part_name = "chiasm"; // lower case part of the name to look for
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) & !s.IsEmpty);
            title = "Optic Chiasm:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Optic Nerve Lt and Rt **********************************************************************************************
            // Getting all the structures containing "optic" and "nerve" 
            part_name = "optic"; // lower case part of the name to look for
            partial_set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) 
                    & s.Id.ToLower().Contains("nerve"));

            // Col AD: Optic Nerve Lt
            set_of_structs = partial_set_of_structs.Where(s => s.Id.ToLower().Contains("l") & !s.IsEmpty);
            title = "Optic Nerve Lt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }
            
            // Col AE: Optic Nerve Rt
            set_of_structs = partial_set_of_structs.Where(s => !s.Id.ToLower().Contains("l") & !s.IsEmpty);
            title = "Optic Nerve Rt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Col AF: Oral Cavity **********************************************************************************************
            // Getting all the structures containing "oral" and "cav" 
            part_name = "oral"; // lower case part of the name to look for
            set_of_structs = my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains(part_name) & s.Id.ToLower().Contains("cav") & !s.IsEmpty);
            title = "Uninv. Oral Cavity:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Parotids ************************************************************************************************
            // Getting all the structures containing "parotid" 
            part_name = "parotid"; // lower case part of the name to look for
            partial_set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name));

            // COl AG-AI: Parotid Lt
            set_of_structs = partial_set_of_structs.Where(s => s.Id.ToLower().Contains("l") & !s.IsEmpty);
            title = "Parotid Lt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean, V30Gy [%], V20Gy [cm3]
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
                // Col AG: D_Mean,
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                // Col AH: V30Gy [%]
                Abs_Dose = new DoseValue(30.0, "Gy");
                Rel_Vol_Metric = ALCC_QM.V_X_report(my_plan, selected_strct, Abs_Dose, VolumePresentation.Relative);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Rel_Vol_Metric, 2).ToString() + ","; // Not use % symbol for Excel
                // Col AI: V20Gy [cm3]
                Abs_Dose = new DoseValue(20.0, "Gy");
                Vol_Metric = ALCC_QM.V_X_report(my_plan, selected_strct, Abs_Dose, VolumePresentation.AbsoluteCm3);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Vol_Metric, 2).ToString() + ",";
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN, NaN, NaN,"; // 3 colums to skip
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // COl AJ-AL: Parotid Rt
            set_of_structs = partial_set_of_structs.Where(s => !s.Id.ToLower().Contains("l"));
            title = "Parotid Rt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean, V30Gy [%], V20Gy [cm3]
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
                // Col AJ: D_Mean,
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                // Col AK: V30Gy [%]
                Abs_Dose = new DoseValue(30.0, "Gy");
                Rel_Vol_Metric = ALCC_QM.V_X_report(my_plan, selected_strct, Abs_Dose, VolumePresentation.Relative);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Rel_Vol_Metric, 2).ToString() + ","; // Not use % symbol for Excel
                // Col AL: V20Gy [cm3]
                Abs_Dose = new DoseValue(20.0, "Gy");
                Vol_Metric = ALCC_QM.V_X_report(my_plan, selected_strct, Abs_Dose, VolumePresentation.AbsoluteCm3);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Vol_Metric, 2).ToString() + ",";
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN, NaN, NaN,"; // 3 colums to skip
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }
            
            // Col AM: Pharyngeal constrictor  ************************************************************************************
            // Getting all the structures containing "pharyn"
            part_name = "pharyn"; // lower case part of the name to look for
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) & !s.IsEmpty);
            title = "Pharyngeal:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Subman Lt and Rt **********************************************************************************************
            // Col AN-AO: Subman is has not metric defined, thus decided for Dmean (as Parotid)
            // Getting all the structures containing "subman" 
            part_name = "subman"; // lower case part of the name to look for
            partial_set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name));

            // COl AN: Subman Lt
            set_of_structs = partial_set_of_structs.Where(s => s.Id.ToLower().Replace("mandibular","").Contains("l") & !s.IsEmpty);
            title = "Subman. Lt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // COl AO: Subman. Rt
            set_of_structs = partial_set_of_structs.Where(s => !s.Id.ToLower().Replace("mandibular", "").Contains("l") & !s.IsEmpty);
            title = "Subman. Rt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }
            
            // Mass Muscle Lt and Rt ******************************************************************************************
            // Getting all the structures containing "mass" and "muscle" 
            part_name = "mass"; // lower case part of the name to look for
            partial_set_of_structs = my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains(part_name) & s.Id.ToLower().Contains("muscle"));

            // COl AP: Mass Muscle Lt
            set_of_structs = partial_set_of_structs.Where(s => !s.Id.ToLower().Replace("eter", "").Contains("r") & !s.IsEmpty);
            title = "Masseter Muscle Lt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,";
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // COl AQ: Mass Muscle Rt  
            set_of_structs = partial_set_of_structs.Where(s => s.Id.ToLower().Replace("eter", "").Contains("r") & !s.IsEmpty);
            title = "Mass. Muscle Rt:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Mean
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ", "; 
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN,"; 
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }

            // Col AR: Brachial Plexus  Col AS is empty!!********************************************************************
            // Getting all the structures containing "brach" 
            part_name = "brach"; // lower case part of the name to look for
            set_of_structs = my_plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains(part_name) & !s.IsEmpty);
            title = "Brachial Plexus:";
            if (set_of_structs.Any())
            {
                if (set_of_structs.Count() > 1)
                {
                    selectOneStruct = new SelectOneStruct(title, my_plan, set_of_structs);
                    selected_strct = selectOneStruct.Get_Selected();
                }
                else
                {
                    selected_strct = set_of_structs.First();
                }
                // D_Max
                Dose_Metric = ALCC_QM.Max_Dose(my_plan, selected_strct, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ", ,";// 2 comas for Col AS!
                selected_structs.Add(Tuple.Create(title + " - - - - ", selected_strct.Id));
            }
            else
            {
                VPSRG_HN_track = VPSRG_HN_track + "NaN, ,"; // 2 comas for Col AS!
                selected_structs.Add(Tuple.Create(title + " - - - - ", "None"));
            }


            // PTVs Sorting them by dose written in name
            List<Tuple<Structure, int>> ptv_mindose=new List<Tuple<Structure, int>>();
            set_of_ptvs = my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains("ptv") & !s.Id.ToLower().Contains("ip") & !s.IsEmpty);
            foreach (Structure ptv in set_of_ptvs)
            {
                String text_dose = ptv.Id;
                text_dose = new string(text_dose.Where(x => char.IsDigit(x)).ToArray());
                if (int.TryParse(text_dose, out int dose))
                { }
                else { dose = 0; }
                ptv_mindose.Add(Tuple.Create(ptv,dose));
            }
            var sorted_ptvs = ptv_mindose.OrderByDescending(x => x.Item2).ToList();

            double dose_ptv_high = new double(); // for keeping dose value from name

            // PTV high
            // D_2% [Gy], V95% [%], D_Mean [Gy]
            if (sorted_ptvs.Any() && !sorted_ptvs.First().Item1.IsEmpty)
            { // If there is any, then get the required metrics from first

                //*************** Select PTV High structure
                my_list.Clear();
                foreach (Structure ptv in my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains("ptv") & 
                    s.Id.ToLower().Contains(sorted_ptvs.First().Item2.ToString()) 
                    & !s.IsEmpty))
                { my_list.Add(ptv.Id); }
                selectDiag = new SelectBox(my_list, "PTV High");
                selected = selectDiag.Get_Item();
                Structure PTV_high = my_plan.StructureSet.Structures.Where(s=> s.Id.Equals(selected)).First();
                selected_structs.Add(Tuple.Create("PTV High:    ", PTV_high.Id));
                
                // Col AT: D_2% [Gy]
                Rel_Vol_Metric = 2.0; // 2% of Vol
                Dose_Metric = ALCC_QM.D_X_report(my_plan, PTV_high,
                    Rel_Vol_Metric, VolumePresentation.Relative, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";
                
                // Col AU: V95% [%]
                Rel_Dose = new DoseValue(95.0, "%");
                Rel_Vol_Metric = ALCC_QM.V_X_report(my_plan, PTV_high,
                    Rel_Dose, VolumePresentation.Relative);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Rel_Vol_Metric, 2).ToString() + ","; // Not use % symbol for Excel

                // Col AV: D_Mean [Gy]
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, sorted_ptvs.First().Item1, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";

                // Delete first element from sorted_ptvs as already used
                dose_ptv_high = sorted_ptvs.First().Item2;// keep dose value from name
                sorted_ptvs.Remove(sorted_ptvs.First());
            }
            else
            {   // 9 colums to skip!! and state 0 targets!
                VPSRG_HN_track = VPSRG_HN_track + "NaN, NaN, NaN, NaN, NaN, NaN, NaN, , 0";
                selected_structs.Add(Tuple.Create("PTV High:    ", "None"));
                // and tht is it!, nothing more to record!
            }

            // PTV Int
            // See if 2 ptvs left
            if (sorted_ptvs.Count==2)
            {
                //*************** Select PTV Int structure
                my_list.Clear();
                foreach (Structure ptv in my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains("ptv") &
                    s.Id.ToLower().Contains(sorted_ptvs.First().Item2.ToString())
                    & !s.IsEmpty))
                { my_list.Add(ptv.Id); }
                selectDiag = new SelectBox(my_list, "PTV Int");
                selected = selectDiag.Get_Item();
                Structure PTV_int = my_plan.StructureSet.Structures.Where(s => s.Id.Equals(selected)).First();
                selected_structs.Add(Tuple.Create("PTV Int. ID:    ", PTV_int.Id));
                double dose_ptv_int = sorted_ptvs.First().Item2;
                // Have still 2 ptvs, process here PTV Int
                // Col AW: V95% [%]  95% of its own dose thus 95%*dose_ptv_int/dose_ptv_high
                Rel_Dose = new DoseValue(95.0* dose_ptv_int / dose_ptv_high, "%");
                Rel_Vol_Metric = ALCC_QM.V_X_report(my_plan, PTV_int,
                    Rel_Dose, VolumePresentation.Relative);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Rel_Vol_Metric, 2).ToString() + ","; // Not use % symbol for Excel

                // Col AX: D_Mean [Gy]
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, PTV_int, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";

                // Delete first element from sorted_ptvs as already used
                sorted_ptvs.Remove(sorted_ptvs.First());

                //*************** Select PTV Low structure
                my_list.Clear();
                foreach (Structure ptv in my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains("ptv") &
                    s.Id.ToLower().Contains(sorted_ptvs.First().Item2.ToString())
                    & !s.IsEmpty))
                { my_list.Add(ptv.Id); }
                selectDiag = new SelectBox(my_list, "PTV Low");
                selected = selectDiag.Get_Item();
                Structure PTV_low = my_plan.StructureSet.Structures.Where(s => s.Id.Equals(selected)).First();
                selected_structs.Add(Tuple.Create("PTV Low ID:    ", PTV_low.Id));
                double dose_ptv_low = sorted_ptvs.First().Item2;
                // Have still 1 ptv left: PTV Low
                // Col AY: V95% [%] 95% of its own dose thus 95%*dose_ptv_low/dose_ptv_high
                Rel_Dose = new DoseValue(95.0* dose_ptv_low / dose_ptv_high, "%");
                Rel_Vol_Metric = ALCC_QM.V_X_report(my_plan, PTV_low,
                    Rel_Dose, VolumePresentation.Relative);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Rel_Vol_Metric, 2).ToString() + ","; // Not use % symbol for Excel

                // Col AZ: D_Mean [Gy]
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, PTV_low, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";

                // Delete first element from sorted_ptvs as already used
                sorted_ptvs.Remove(sorted_ptvs.First()); // Needed to skip following if which process a PTV Low 
                                                         // when there is no PTV int   

                // Col BB:Skip Col BA and write # of targets
                VPSRG_HN_track = VPSRG_HN_track +  ", 3"; // and that is it!
            }

            // PTV Low Without PTV Int
            if (sorted_ptvs.Count == 1)
            {
                selected_structs.Add(Tuple.Create("PTV Int. ID:    ", "None"));
                //*************** Select PTV Low structure
                my_list.Clear();
                foreach (Structure ptv in my_plan.StructureSet.Structures.
                    Where(s => s.Id.ToLower().Contains("ptv") &
                    s.Id.ToLower().Contains(sorted_ptvs.First().Item2.ToString())
                    & !s.IsEmpty))
                { my_list.Add(ptv.Id); }
                selectDiag = new SelectBox(my_list, "PTV Low");
                selected = selectDiag.Get_Item();
                Structure PTV_low = my_plan.StructureSet.Structures.Where(s => s.Id.Equals(selected)).First();
                selected_structs.Add(Tuple.Create("PTV Low ID:    ", PTV_low.Id));
                double dose_ptv_low = sorted_ptvs.First().Item2;
                
                // Skip 2 Col AW-AX as no PTV Int
                VPSRG_HN_track = VPSRG_HN_track + ", ,";

                // Have still 1 ptv left: PTV Low
                // Col AY: V95% [%] 95% of its own dose thus 95%*dose_ptv_low/dose_ptv_high
                Rel_Dose = new DoseValue(95.0 * dose_ptv_low / dose_ptv_high, "%");
                Rel_Vol_Metric = ALCC_QM.V_X_report(my_plan, PTV_low,
                    Rel_Dose, VolumePresentation.Relative);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Rel_Vol_Metric, 2).ToString() + "%,";

                // Col AZ: D_Mean [Gy]
                Dose_Metric = ALCC_QM.Mean_Dose(my_plan, PTV_low, DoseValuePresentation.Absolute);
                VPSRG_HN_track = VPSRG_HN_track + Math.Round(Dose_Metric.Dose, 2).ToString() + ",";

                // Col BB:Skip Col BA and write # of targets
                VPSRG_HN_track = VPSRG_HN_track + ", 2"; // and that is it!

            }

            String file_name = @"c:\temp\VPSRG_HN_" + my_patient.Id + ".txt";
                // create or overwrite
                System.IO.File.WriteAllText(file_name, VPSRG_HN_track, Encoding.UTF8);


            // Build output text
            String text = null;
            foreach (var item in selected_structs)
            {
                text = text + item.Item1 + item.Item2 + System.Environment.NewLine;
            }
            System.Windows.MessageBox.Show(text);

            String file_name_str = @"c:\temp\VPSRG_HN_" + my_patient.Id + "_SelectedStructures.txt";
            // create or overwrite
            System.IO.File.WriteAllText(file_name_str, text, Encoding.UTF8);

            System.Windows.MessageBox.Show("File " + file_name + " saved." + System.Environment.NewLine +
                " (path copied to clipboard)");
            System.Windows.Clipboard.Clear();
            System.Windows.Clipboard.SetText(@"c:\temp\");

        }

        

    }



    



}





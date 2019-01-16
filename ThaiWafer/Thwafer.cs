using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using SpssLib.SpssDataset;
using SpssLib.DataReader;
using System.Data.SqlClient;

namespace ThaiWafer
{
    class Thwafer
    {
        static void Main(string[] args)
        {
            int SurveyID = 600550;
            string SurveyPeriod = "2018-09-30";//Survey Period
            string Country = "THAILAND";//survey country
            ThaiInsertion iobj = new ThaiInsertion();
            string questions = "ID,Quota_Gender,E3,Quota_Location,Quota_Age,E2,E1,S4.1_BKK,S4.2_BKK,S4.1_UPC,S4.2_UPC,Wave,S5_1,S5_2,S5_3,S5_4,S5_5,S5_6,S5_7,S5_8,S5_9,S5_10,S5_11,S5_12,S5_13,S5_14,S5_15,S5_16,S5_17,S5_18,S5_19,S6,S7_1,S7_2,S7_3,S7_4,S7_5,S7_6,S7_7,S7_8,S7_9,S7_10,S7_11,S7_12,S7_13,S7_14,S7_15,S7_16,S7_17,S7_18,S7_19,S8,S10.1,S11.1,r1.1,r1.2_2,r1.2_6,r1.2_7,r1.2_8,r1.2_9,r1.2_14,r1.2_3,r1.2_4,r1.2_11,r1.2_12,r1.2_13,r1.2_102,r3_2,r3_6,r3_7,r3_8,r3_9,r3_14,r3_3,r3_4,r3_11,r3_12,r3_13,r3_102,r2.1,r2.2_2,r2.2_6,r2.2_7,r2.2_8,r2.2_9,r2.2_14,r2.2_3,r2.2_4,r2.2_11,r2.2_12,r2.2_13,r2.2_102,r4_2,r4_6,r4_7,r4_8,r4_9,r4_14,r4_3,r4_4,r4_11,r4_12,r4_13,r4_102,r5.1_2,r5.1_6,r5.1_7,r5.1_8,r5.1_9,r5.1_14,r5.1_3,r5.1_4,r5.1_11,r5.1_12,r5.1_13,r5.1_102,r5.2_2,r5.2_6,r5.2_7,r5.2_8,r5.2_9,r5.2_14,r5.2_3,r5.2_4,r5.2_11,r5.2_12,r5.2_13,r5.2_102,r6.1_2,r6.1_6,r6.1_7,r6.1_8,r6.1_9,r6.1_14,r6.1_3,r6.1_4,r6.1_11,r6.1_12,r6.1_13,r6.1_102,r6.2_2,r6.2_6,r6.2_7,r6.2_8,r6.2_9,r6.2_14,r6.2_3,r6.2_4,r6.2_11,r6.2_12,r6.2_13,r6.2_102,r5.3,r6.3,r6.4_2,r6.4_6,r6.4_7,r6.4_8,r6.4_9,r6.4_14,r6.4_3,r6.4_4,r6.4_11,r6.4_12,r6.4_13,r6.4_102,r6.5,r1.1_net,r1.2_net_1,r1.2_net_2,r1.2_net_5,r1.2_net_3,r1.2_net_4,r3_net_1,r3_net_2,r3_net_5,r3_net_3,r3_net_4,r2.1_net,r2.2_net_1,r2.2_net_2,r2.2_net_5,r2.2_net_3,r2.2_net_4,r4_net_1,r4_net_2,r4_net_5,r4_net_3,r4_net_4,r5.1_net_1,r5.1_net_2,r5.1_net_5,r5.1_net_3,r5.1_net_4,r5.2_net_1,r5.2_net_2,r5.2_net_5,r5.2_net_3,r5.2_net_4,r6.1_net_1,r6.1_net_2,r6.1_net_5,r6.1_net_3,r6.1_net_4,r6.2_net_1,r6.2_net_2,r6.2_net_5,r6.2_net_3,r6.2_net_4,r5.3_net,r6.3_net,r6.4_net_1,r6.4_net_2,r6.4_net_5,r6.4_net_3,r6.4_net_4,r6.5_net,r14h1_1,r14h1_2,r14h1_3,r14h1_4,r14h1_5,r14h1_6,r14h1_7,r14h1_8,r14h1_9,r14h1_10,r14h1_12,r14h1_13,r14h1_14,r14h1_15,r14h1_17,r14h2_1,r14h2_2,r14h2_3,r14h2_4,r14h2_5,r14h2_6,r14h2_7,r14h2_8,r14h2_9,r14h2_10,r14h2_12,r14h2_13,r14h2_14,r14h2_15,r14h2_17,r14h3_1,r14h3_2,r14h3_3,r14h3_4,r14h3_5,r14h3_6,r14h3_7,r14h3_8,r14h3_9,r14h3_10,r14h3_12,r14h3_13,r14h3_14,r14h3_15,r14h3_17,r14h4_1,r14h4_2,r14h4_3,r14h4_4,r14h4_5,r14h4_6,r14h4_7,r14h4_8,r14h4_9,r14h4_10,r14h4_12,r14h4_13,r14h4_14,r14h4_15,r14h4_17";// dashboard variable value

            string[] spss_variable_name = questions.Split(',');
            SpssReader spssDataset;
            using (FileStream fileStream = new FileStream(@"D:\spssparsing\Thailand\Thai-Sep18.sav", FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            {
                spssDataset = new SpssReader(fileStream); // Create the reader, this will read the file header
                foreach (string spss_v in spss_variable_name)
                {
                    foreach (var variable in spssDataset.Variables)  // Iterate through all the varaibles
                    {
                        if (variable.Name.Equals(spss_v))
                        {
                            foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                            {
                                string VARIABLE_NAME = spss_v;
                                string VARIABLE_NAME_SUB_NAME = variable.Name;
                                string VARIABLE_ID = label.Key.ToString();
                                string VARIABLE_VALUE = label.Value;
                                string VARIABLE_NAME_QUESTION = variable.Label;
                                string BASE_VARIABLE_NAME = variable.Name;
                                // iobj.insert_Dashboard_variable_values(VARIABLE_NAME, VARIABLE_NAME_SUB_NAME, VARIABLE_ID, VARIABLE_VALUE, VARIABLE_NAME_QUESTION, SurveyID, Country, BASE_VARIABLE_NAME, SurveyPeriod);
                            }
                        }

                    }
                }
                foreach (var record in spssDataset.Records)
                {
                    string UserID = null;
                    string variable_name;
                    string u_id = null;
                    decimal Weight = 1;
                    string AttendedOn = "-- Not Available --";
                    string Gender = "-- Not Available --";
                    string MaritalStatus = "-- Not Available --";
                    string Location = "-- Not Available --";
                    string AgeGroup = "-- Not Available --";
                    string Education = "-- Not Available --";
                    string Occupation = "-- Not Available --";
                    string PersonalInc_BKK = "-- Not Available --";
                    string HHIncome_BKK = "-- Not Available --";
                    string PersonalInc_UPC = "-- Not Available --";
                    string HHIncome_UPC = "-- Not Available --";
                    string Period = "-- Not Available --";
                    string S5_1 = "-- Not Available --";
                    string S5_2 = "-- Not Available --";
                    string S5_3 = "-- Not Available --";
                    string S5_4 = "-- Not Available --";
                    string S5_5 = "-- Not Available --";
                    string S5_6 = "-- Not Available --";
                    string S5_7 = "-- Not Available --";
                    string S5_8 = "-- Not Available --";
                    string S5_9 = "-- Not Available --";
                    string S5_10 = "-- Not Available --";
                    string S5_11 = "-- Not Available --";
                    string S5_12 = "-- Not Available --";
                    string S5_13 = "-- Not Available --";
                    string S5_14 = "-- Not Available --";
                    string S5_15 = "-- Not Available --";
                    string S5_16 = "-- Not Available --";
                    string S5_17 = "-- Not Available --";
                    string S5_18 = "-- Not Available --";
                    string S5_19 = "-- Not Available --";
                    string S6 = "-- Not Available --";
                    string S7_1 = "-- Not Available --";
                    string S7_2 = "-- Not Available --";
                    string S7_3 = "-- Not Available --";
                    string S7_4 = "-- Not Available --";
                    string S7_5 = "-- Not Available --";
                    string S7_6 = "-- Not Available --";
                    string S7_7 = "-- Not Available --";
                    string S7_8 = "-- Not Available --";
                    string S7_9 = "-- Not Available --";
                    string S7_10 = "-- Not Available --";
                    string S7_11 = "-- Not Available --";
                    string S7_12 = "-- Not Available --";
                    string S7_13 = "-- Not Available --";
                    string S7_14 = "-- Not Available --";
                    string S7_15 = "-- Not Available --";
                    string S7_16 = "-- Not Available --";
                    string S7_17 = "-- Not Available --";
                    string S7_18 = "-- Not Available --";
                    string S7_19 = "-- Not Available --";
                    string S8 = "-- Not Available --";
                    string S10_1 = "-- Not Available --";
                    string S11_1 = "-- Not Available --";
                    string BrTom = "-- Not Available --";
                    string BrSpont_BBengRed = "-- Not Available --";
                    string BrSpont_TivWafer = "-- Not Available --";
                    string BrSpont_TivTwin = "-- Not Available --";
                    string BrSpont_TivCombo = "-- Not Available --";
                    string BrSpont_TivJumbo = "-- Not Available --";
                    string BrSpont_Sanghai = "-- Not Available --";
                    string BrSpont_BBengHazel = "-- Not Available --";
                    string BrSpont_BBengMaxx = "-- Not Available --";
                    string BrSpont_VoizWaffle = "-- Not Available --";
                    string BrSpont_Voizwafflmini = "-- Not Available --";
                    string BrSpont_Bissin = "-- Not Available --";
                    string BrSpont_CalChees = "-- Not Available --";
                    string BrAid_BBengRed = "-- Not Available --";
                    string BrAid_TivWafer = "-- Not Available --";
                    string BrAid_TivTwin = "-- Not Available --";
                    string BrAid_TivCombo = "-- Not Available --";
                    string BrAid_TivJumbo = "-- Not Available --";
                    string BrAid_Sanghai = "-- Not Available --";
                    string BrAid_BBengHazel = "-- Not Available --";
                    string BrAid_BBengMaxx = "-- Not Available --";
                    string BrAid_VoizWaffle = "-- Not Available --";
                    string BrAid_Voizwafflmini = "-- Not Available --";
                    string BrAid_Bissin = "-- Not Available --";
                    string BrAid_CalChees = "-- Not Available --";
                    string AdTom = "-- Not Available --";
                    string AdSpont_BBengRed = "-- Not Available --";
                    string AdSpont_TivWafer = "-- Not Available --";
                    string AdSpont_TivTwin = "-- Not Available --";
                    string AdSpont_TivCombo = "-- Not Available --";
                    string AdSpont_TivJumbo = "-- Not Available --";
                    string AdSpont_Sanghai = "-- Not Available --";
                    string AdSpont_BBengHazel = "-- Not Available --";
                    string AdSpont_BBengMaxx = "-- Not Available --";
                    string AdSpont_VoizWaffle = "-- Not Available --";
                    string AdSpont_Voizwafflmini = "-- Not Available --";
                    string AdSpont_Bissin = "-- Not Available --";
                    string AdSpont_CalChees = "-- Not Available --";
                    string AdAid_BBengRed = "-- Not Available --";
                    string AdAid_TivWafer = "-- Not Available --";
                    string AdAid_TivTwin = "-- Not Available --";
                    string AdAid_TivCombo = "-- Not Available --";
                    string AdAid_TivJumbo = "-- Not Available --";
                    string AdAid_Sanghai = "-- Not Available --";
                    string AdAid_BBengHazel = "-- Not Available --";
                    string AdAid_BBengMaxx = "-- Not Available --";
                    string AdAid_VoizWaffle = "-- Not Available --";
                    string AdAid_Voizwafflmini = "-- Not Available --";
                    string AdAid_Bissin = "-- Not Available --";
                    string AdAided_CalChees = "-- Not Available --";
                    string EverCons_BBengRed = "-- Not Available --";
                    string EverCons_TivWafer = "-- Not Available --";
                    string EverCons_TivTwin = "-- Not Available --";
                    string EverCons_TivCombo = "-- Not Available --";
                    string EverCons_TivJumbo = "-- Not Available --";
                    string EverCons_Sanghai = "-- Not Available --";
                    string EverCons_BBengHazel = "-- Not Available --";
                    string EverCons_BBengMaxx = "-- Not Available --";
                    string EverCons_VoizWaffle = "-- Not Available --";
                    string EverCons_Voizwafflmini = "-- Not Available --";
                    string EverCons_Bissin = "-- Not Available --";
                    string EverCons_CalChees = "-- Not Available --";
                    string ConsL3M_BBengRed = "-- Not Available --";
                    string ConsL3M_TivWafer = "-- Not Available --";
                    string ConsL3M_TivTwin = "-- Not Available --";
                    string ConsL3M_TivCombo = "-- Not Available --";
                    string ConsL3M_TivJumbo = "-- Not Available --";
                    string ConsL3M_Sanghai = "-- Not Available --";
                    string ConsL3M_BBengHazel = "-- Not Available --";
                    string ConsL3M_BBengMaxx = "-- Not Available --";
                    string ConsL3M_VoizWaffle = "-- Not Available --";
                    string ConsL3M_Voizwafflmini = "-- Not Available --";
                    string ConsL3M_Bissin = "-- Not Available --";
                    string ConsL3M_CalChees = "-- Not Available --";
                    string ConsL1M_BBengRed = "-- Not Available --";
                    string ConsL1M_TivWafer = "-- Not Available --";
                    string ConsL1M_TivTwin = "-- Not Available --";
                    string ConsL1M_TivCombo = "-- Not Available --";
                    string ConsL1M_TivJumbo = "-- Not Available --";
                    string ConsL1M_Sanghai = "-- Not Available --";
                    string ConsL1M_BBengHazel = "-- Not Available --";
                    string ConsL1M_BBengMaxx = "-- Not Available --";
                    string ConsL1M_VoizWaffle = "-- Not Available --";
                    string ConsL1M_Voizwafflmini = "-- Not Available --";
                    string ConsL1M_Bissin = "-- Not Available --";
                    string ConsL1M_CalChees = "-- Not Available --";
                    string ConsL1W_BBengRed = "-- Not Available --";
                    string ConsL1W_TivWafer = "-- Not Available --";
                    string ConsL1W_TivTwin = "-- Not Available --";
                    string ConsL1W_TivCombo = "-- Not Available --";
                    string ConsL1W_TivJumbo = "-- Not Available --";
                    string ConsL1W_Sanghai = "-- Not Available --";
                    string ConsL1W_BBengHazel = "-- Not Available --";
                    string ConsL1W_BBengMaxx = "-- Not Available --";
                    string ConsL1W_VoizWaffle = "-- Not Available --";
                    string ConsL1W_Voizwafflmini = "-- Not Available --";
                    string ConsL1W_Bissin = "-- Not Available --";
                    string ConsL1W_CalChees = "-- Not Available --";
                    string Bumo = "-- Not Available --";
                    string PreBumo = "-- Not Available --";
                    string FavBrand_BBengRed = "-- Not Available --";
                    string FavBrand_TivWafer = "-- Not Available --";
                    string FavBrand_TivTwin = "-- Not Available --";
                    string FavBrand_TivCombo = "-- Not Available --";
                    string FavBrand_TivJumbo = "-- Not Available --";
                    string FavBrand_Sanghai = "-- Not Available --";
                    string FavBrand_BBengHazel = "-- Not Available --";
                    string FavBrand_BBengMaxx = "-- Not Available --";
                    string FavBrand_VoizWaffle = "-- Not Available --";
                    string FavBrand_Voizwafflmini = "-- Not Available --";
                    string FavBrand_Bissin = "-- Not Available --";
                    string FavBrand_CalChees = "-- Not Available --";
                    string MostFavBrand = "-- Not Available --";
                    string NetBrTom = "-- Not Available --";
                    string BrSpontNet_BBengNet = "-- Not Available --";
                    string BrSpontNet_TivoliNet = "-- Not Available --";
                    string BrSpontNet_SanghaiNet = "-- Not Available --";
                    string BrSpontNet_VoizNet = "-- Not Available --";
                    string BrSpontNet_BissinNet = "-- Not Available --";
                    string BrAidNet_BBengNet = "-- Not Available --";
                    string BrAidNet_TivoliNet = "-- Not Available --";
                    string BrAidNet_SanghaiNet = "-- Not Available --";
                    string BrAidNet_VoizNet = "-- Not Available --";
                    string BrAidNet_BissinNet = "-- Not Available --";
                    string NetAdTom = "-- Not Available --";
                    string AdSpontNet_BBengNet = "-- Not Available --";
                    string AdSpontNet_TivoliNet = "-- Not Available --";
                    string AdSpontNet_SanghaiNet = "-- Not Available --";
                    string AdSpontNet_VoizNet = "-- Not Available --";
                    string AdSpontNet_BissinNet = "-- Not Available --";
                    string AdAidNet_BBengNet = "-- Not Available --";
                    string AdAidNet_TivoliNet = "-- Not Available --";
                    string AdAidNet_SanghaiNet = "-- Not Available --";
                    string AdAidNet_VoizNet = "-- Not Available --";
                    string AdAidNet_BissinNet = "-- Not Available --";
                    string EverConsNet_BBengNet = "-- Not Available --";
                    string EverConsNet_TivoliNet = "-- Not Available --";
                    string EverConsNet_SanghaiNet = "-- Not Available --";
                    string EverConsNet_VoizNet = "-- Not Available --";
                    string EverConsNet_BissinNet = "-- Not Available --";
                    string ConsL3MNet_BBengNet = "-- Not Available --";
                    string ConsL3MNet_TivoliNet = "-- Not Available --";
                    string ConsL3MNet_SanghaiNet = "-- Not Available --";
                    string ConsL3MNet_VoizNet = "-- Not Available --";
                    string ConsL3MNet_BissinNet = "-- Not Available --";
                    string ConsL1MNet_BBengNet = "-- Not Available --";
                    string ConsL1MNet_TivoliNet = "-- Not Available --";
                    string ConsL1MNet_SanghaiNet = "-- Not Available --";
                    string ConsL1MNet_VoizNet = "-- Not Available --";
                    string ConsL1MNet_BissinNet = "-- Not Available --";
                    string ConsL1WNet_BBengNet = "-- Not Available --";
                    string ConsL1WNet_TivoliNet = "-- Not Available --";
                    string ConsL1WNet_SanghaiNet = "-- Not Available --";
                    string ConsL1WNet_VoizNet = "-- Not Available --";
                    string ConsL1WNet_BissinNet = "-- Not Available --";
                    string Bumo_Net = "-- Not Available --";
                    string PreBumo_Net = "-- Not Available --";
                    string FavBrandNet_BBengNet = "-- Not Available --";
                    string FavBrandNet_TivoliNet = "-- Not Available --";
                    string FavBrandNet_SanghaiNet = "-- Not Available --";
                    string FavBrandNet_VoizNet = "-- Not Available --";
                    string FavBrandNet_BissinNet = "-- Not Available --";
                    string MostFavBrand_Net = "-- Not Available --";
                    string r14h1_1 = "-- Not Available --";
                    string r14h1_2 = "-- Not Available --";
                    string r14h1_3 = "-- Not Available --";
                    string r14h1_4 = "-- Not Available --";
                    string r14h1_5 = "-- Not Available --";
                    string r14h1_6 = "-- Not Available --";
                    string r14h1_7 = "-- Not Available --";
                    string r14h1_8 = "-- Not Available --";
                    string r14h1_9 = "-- Not Available --";
                    string r14h1_10 = "-- Not Available --";
                    string r14h1_12 = "-- Not Available --";
                    string r14h1_13 = "-- Not Available --";
                    string r14h1_14 = "-- Not Available --";
                    string r14h1_15 = "-- Not Available --";
                    string r14h1_17 = "-- Not Available --";
                    string r14h2_1 = "-- Not Available --";
                    string r14h2_2 = "-- Not Available --";
                    string r14h2_3 = "-- Not Available --";
                    string r14h2_4 = "-- Not Available --";
                    string r14h2_5 = "-- Not Available --";
                    string r14h2_6 = "-- Not Available --";
                    string r14h2_7 = "-- Not Available --";
                    string r14h2_8 = "-- Not Available --";
                    string r14h2_9 = "-- Not Available --";
                    string r14h2_10 = "-- Not Available --";
                    string r14h2_12 = "-- Not Available --";
                    string r14h2_13 = "-- Not Available --";
                    string r14h2_14 = "-- Not Available --";
                    string r14h2_15 = "-- Not Available --";
                    string r14h2_17 = "-- Not Available --";
                    string r14h3_1 = "-- Not Available --";
                    string r14h3_2 = "-- Not Available --";
                    string r14h3_3 = "-- Not Available --";
                    string r14h3_4 = "-- Not Available --";
                    string r14h3_5 = "-- Not Available --";
                    string r14h3_6 = "-- Not Available --";
                    string r14h3_7 = "-- Not Available --";
                    string r14h3_8 = "-- Not Available --";
                    string r14h3_9 = "-- Not Available --";
                    string r14h3_10 = "-- Not Available --";
                    string r14h3_12 = "-- Not Available --";
                    string r14h3_13 = "-- Not Available --";
                    string r14h3_14 = "-- Not Available --";
                    string r14h3_15 = "-- Not Available --";
                    string r14h3_17 = "-- Not Available --";
                    string r14h4_1 = "-- Not Available --";
                    string r14h4_2 = "-- Not Available --";
                    string r14h4_3 = "-- Not Available --";
                    string r14h4_4 = "-- Not Available --";
                    string r14h4_5 = "-- Not Available --";
                    string r14h4_6 = "-- Not Available --";
                    string r14h4_7 = "-- Not Available --";
                    string r14h4_8 = "-- Not Available --";
                    string r14h4_9 = "-- Not Available --";
                    string r14h4_10 = "-- Not Available --";
                    string r14h4_12 = "-- Not Available --";
                    string r14h4_13 = "-- Not Available --";
                    string r14h4_14 = "-- Not Available --";
                    string r14h4_15 = "-- Not Available --";
                    string r14h4_17 = "-- Not Available --";


                    foreach (var variable in spssDataset.Variables)
                    {
                        foreach (string spss_v in spss_variable_name)
                        {
                            if (variable.Name.Equals(spss_v))
                            {
                                variable_name = variable.Name;
                                switch (variable_name)
                                {
                                    case "ID":
                                        {
                                            u_id = Convert.ToString(record.GetValue(variable));
                                            UserID = find_UserID(SurveyID, SurveyPeriod, u_id);
                                            //userID = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    
                                    case "Quota_Gender":
                                        {
                                            Gender = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "E3":
                                        {
                                            MaritalStatus = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "Quota_Location":
                                        {
                                            Location = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "Quota_Age":
                                        {
                                            AgeGroup = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "E2":
                                        {
                                            Education = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "E1":
                                        {
                                            Occupation = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.1_BKK":
                                        {
                                            PersonalInc_BKK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.2_BKK":
                                        {
                                            HHIncome_BKK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.1_UPC":
                                        {
                                            PersonalInc_UPC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.2_UPC":
                                        {
                                            HHIncome_UPC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "Wave":
                                        {
                                            Period = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_1":
                                        {
                                            S5_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_2":
                                        {
                                            S5_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_3":
                                        {
                                            S5_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_4":
                                        {
                                            S5_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_5":
                                        {
                                            S5_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_6":
                                        {
                                            S5_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_7":
                                        {
                                            S5_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_8":
                                        {
                                            S5_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_9":
                                        {
                                            S5_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_10":
                                        {
                                            S5_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_11":
                                        {
                                            S5_11 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_12":
                                        {
                                            S5_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_13":
                                        {
                                            S5_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_14":
                                        {
                                            S5_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_15":
                                        {
                                            S5_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_16":
                                        {
                                            S5_16 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_17":
                                        {
                                            S5_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_18":
                                        {
                                            S5_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_19":
                                        {
                                            S5_19 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S6":
                                        {
                                            S6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_1":
                                        {
                                            S7_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_2":
                                        {
                                            S7_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_3":
                                        {
                                            S7_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_4":
                                        {
                                            S7_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_5":
                                        {
                                            S7_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_6":
                                        {
                                            S7_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_7":
                                        {
                                            S7_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_8":
                                        {
                                            S7_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_9":
                                        {
                                            S7_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_10":
                                        {
                                            S7_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_11":
                                        {
                                            S7_11 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_12":
                                        {
                                            S7_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_13":
                                        {
                                            S7_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_14":
                                        {
                                            S7_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_15":
                                        {
                                            S7_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_16":
                                        {
                                            S7_16 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_17":
                                        {
                                            S7_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_18":
                                        {
                                            S7_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_19":
                                        {
                                            S7_19 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S8":
                                        {
                                            S8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S10.1":
                                        {
                                            S10_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S11.1":
                                        {
                                            S11_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.1":
                                        {
                                            BrTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_2":
                                        {
                                            BrSpont_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_6":
                                        {
                                            BrSpont_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_7":
                                        {
                                            BrSpont_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_8":
                                        {
                                            BrSpont_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_9":
                                        {
                                            BrSpont_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_14":
                                        {
                                            BrSpont_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_3":
                                        {
                                            BrSpont_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_4":
                                        {
                                            BrSpont_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_11":
                                        {
                                            BrSpont_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_12":
                                        {
                                            BrSpont_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_13":
                                        {
                                            BrSpont_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_102":
                                        {
                                            BrSpont_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_2":
                                        {
                                            BrAid_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_6":
                                        {
                                            BrAid_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_7":
                                        {
                                            BrAid_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_8":
                                        {
                                            BrAid_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_9":
                                        {
                                            BrAid_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_14":
                                        {
                                            BrAid_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_3":
                                        {
                                            BrAid_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_4":
                                        {
                                            BrAid_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_11":
                                        {
                                            BrAid_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_12":
                                        {
                                            BrAid_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_13":
                                        {
                                            BrAid_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_102":
                                        {
                                            BrAid_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.1":
                                        {
                                            AdTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_2":
                                        {
                                            AdSpont_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_6":
                                        {
                                            AdSpont_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_7":
                                        {
                                            AdSpont_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_8":
                                        {
                                            AdSpont_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_9":
                                        {
                                            AdSpont_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_14":
                                        {
                                            AdSpont_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_3":
                                        {
                                            AdSpont_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_4":
                                        {
                                            AdSpont_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_11":
                                        {
                                            AdSpont_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_12":
                                        {
                                            AdSpont_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_13":
                                        {
                                            AdSpont_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_102":
                                        {
                                            AdSpont_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_2":
                                        {
                                            AdAid_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_6":
                                        {
                                            AdAid_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_7":
                                        {
                                            AdAid_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_8":
                                        {
                                            AdAid_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_9":
                                        {
                                            AdAid_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_14":
                                        {
                                            AdAid_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_3":
                                        {
                                            AdAid_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_4":
                                        {
                                            AdAid_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_11":
                                        {
                                            AdAid_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_12":
                                        {
                                            AdAid_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_13":
                                        {
                                            AdAid_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_102":
                                        {
                                            AdAided_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_2":
                                        {
                                            EverCons_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_6":
                                        {
                                            EverCons_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_7":
                                        {
                                            EverCons_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_8":
                                        {
                                            EverCons_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_9":
                                        {
                                            EverCons_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_14":
                                        {
                                            EverCons_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_3":
                                        {
                                            EverCons_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_4":
                                        {
                                            EverCons_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_11":
                                        {
                                            EverCons_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_12":
                                        {
                                            EverCons_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_13":
                                        {
                                            EverCons_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_102":
                                        {
                                            EverCons_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_2":
                                        {
                                            ConsL3M_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_6":
                                        {
                                            ConsL3M_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_7":
                                        {
                                            ConsL3M_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_8":
                                        {
                                            ConsL3M_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_9":
                                        {
                                            ConsL3M_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_14":
                                        {
                                            ConsL3M_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_3":
                                        {
                                            ConsL3M_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_4":
                                        {
                                            ConsL3M_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_11":
                                        {
                                            ConsL3M_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_12":
                                        {
                                            ConsL3M_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_13":
                                        {
                                            ConsL3M_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_102":
                                        {
                                            ConsL3M_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_2":
                                        {
                                            ConsL1M_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_6":
                                        {
                                            ConsL1M_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_7":
                                        {
                                            ConsL1M_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_8":
                                        {
                                            ConsL1M_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_9":
                                        {
                                            ConsL1M_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_14":
                                        {
                                            ConsL1M_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_3":
                                        {
                                            ConsL1M_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_4":
                                        {
                                            ConsL1M_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_11":
                                        {
                                            ConsL1M_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_12":
                                        {
                                            ConsL1M_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_13":
                                        {
                                            ConsL1M_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_102":
                                        {
                                            ConsL1M_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_2":
                                        {
                                            ConsL1W_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_6":
                                        {
                                            ConsL1W_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_7":
                                        {
                                            ConsL1W_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_8":
                                        {
                                            ConsL1W_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_9":
                                        {
                                            ConsL1W_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_14":
                                        {
                                            ConsL1W_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_3":
                                        {
                                            ConsL1W_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_4":
                                        {
                                            ConsL1W_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_11":
                                        {
                                            ConsL1W_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_12":
                                        {
                                            ConsL1W_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_13":
                                        {
                                            ConsL1W_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_102":
                                        {
                                            ConsL1W_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.3":
                                        {
                                            Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.3":
                                        {
                                            PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_2":
                                        {
                                            FavBrand_BBengRed = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_6":
                                        {
                                            FavBrand_TivWafer = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_7":
                                        {
                                            FavBrand_TivTwin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_8":
                                        {
                                            FavBrand_TivCombo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_9":
                                        {
                                            FavBrand_TivJumbo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_14":
                                        {
                                            FavBrand_Sanghai = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_3":
                                        {
                                            FavBrand_BBengHazel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_4":
                                        {
                                            FavBrand_BBengMaxx = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_11":
                                        {
                                            FavBrand_VoizWaffle = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_12":
                                        {
                                            FavBrand_Voizwafflmini = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_13":
                                        {
                                            FavBrand_Bissin = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_102":
                                        {
                                            FavBrand_CalChees = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.5":
                                        {
                                            MostFavBrand = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.1_net":
                                        {
                                            NetBrTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_net_1":
                                        {
                                            BrSpontNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_net_2":
                                        {
                                            BrSpontNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_net_5":
                                        {
                                            BrSpontNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_net_3":
                                        {
                                            BrSpontNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r1.2_net_4":
                                        {
                                            BrSpontNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_net_1":
                                        {
                                            BrAidNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_net_2":
                                        {
                                            BrAidNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_net_5":
                                        {
                                            BrAidNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_net_3":
                                        {
                                            BrAidNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r3_net_4":
                                        {
                                            BrAidNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.1_net":
                                        {
                                            NetAdTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_net_1":
                                        {
                                            AdSpontNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_net_2":
                                        {
                                            AdSpontNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_net_5":
                                        {
                                            AdSpontNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_net_3":
                                        {
                                            AdSpontNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r2.2_net_4":
                                        {
                                            AdSpontNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_net_1":
                                        {
                                            AdAidNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_net_2":
                                        {
                                            AdAidNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_net_5":
                                        {
                                            AdAidNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_net_3":
                                        {
                                            AdAidNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r4_net_4":
                                        {
                                            AdAidNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_net_1":
                                        {
                                            EverConsNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_net_2":
                                        {
                                            EverConsNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_net_5":
                                        {
                                            EverConsNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_net_3":
                                        {
                                            EverConsNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.1_net_4":
                                        {
                                            EverConsNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_net_1":
                                        {
                                            ConsL3MNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_net_2":
                                        {
                                            ConsL3MNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_net_5":
                                        {
                                            ConsL3MNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_net_3":
                                        {
                                            ConsL3MNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.2_net_4":
                                        {
                                            ConsL3MNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_net_1":
                                        {
                                            ConsL1MNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_net_2":
                                        {
                                            ConsL1MNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_net_5":
                                        {
                                            ConsL1MNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_net_3":
                                        {
                                            ConsL1MNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.1_net_4":
                                        {
                                            ConsL1MNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_net_1":
                                        {
                                            ConsL1WNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_net_2":
                                        {
                                            ConsL1WNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_net_5":
                                        {
                                            ConsL1WNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_net_3":
                                        {
                                            ConsL1WNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.2_net_4":
                                        {
                                            ConsL1WNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r5.3_net":
                                        {
                                            Bumo_Net = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.3_net":
                                        {
                                            PreBumo_Net = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_net_1":
                                        {
                                            FavBrandNet_BBengNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_net_2":
                                        {
                                            FavBrandNet_TivoliNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_net_5":
                                        {
                                            FavBrandNet_SanghaiNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_net_3":
                                        {
                                            FavBrandNet_VoizNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.4_net_4":
                                        {
                                            FavBrandNet_BissinNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r6.5_net":
                                        {
                                            MostFavBrand_Net = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_1":
                                        {
                                            r14h1_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_2":
                                        {
                                            r14h1_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_3":
                                        {
                                            r14h1_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_4":
                                        {
                                            r14h1_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_5":
                                        {
                                            r14h1_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_6":
                                        {
                                            r14h1_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_7":
                                        {
                                            r14h1_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_8":
                                        {
                                            r14h1_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_9":
                                        {
                                            r14h1_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_10":
                                        {
                                            r14h1_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_12":
                                        {
                                            r14h1_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_13":
                                        {
                                            r14h1_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_14":
                                        {
                                            r14h1_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_15":
                                        {
                                            r14h1_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h1_17":
                                        {
                                            r14h1_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_1":
                                        {
                                            r14h2_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_2":
                                        {
                                            r14h2_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_3":
                                        {
                                            r14h2_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_4":
                                        {
                                            r14h2_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_5":
                                        {
                                            r14h2_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_6":
                                        {
                                            r14h2_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_7":
                                        {
                                            r14h2_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_8":
                                        {
                                            r14h2_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_9":
                                        {
                                            r14h2_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_10":
                                        {
                                            r14h2_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_12":
                                        {
                                            r14h2_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_13":
                                        {
                                            r14h2_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_14":
                                        {
                                            r14h2_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_15":
                                        {
                                            r14h2_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h2_17":
                                        {
                                            r14h2_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_1":
                                        {
                                            r14h3_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_2":
                                        {
                                            r14h3_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_3":
                                        {
                                            r14h3_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_4":
                                        {
                                            r14h3_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_5":
                                        {
                                            r14h3_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_6":
                                        {
                                            r14h3_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_7":
                                        {
                                            r14h3_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_8":
                                        {
                                            r14h3_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_9":
                                        {
                                            r14h3_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_10":
                                        {
                                            r14h3_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_12":
                                        {
                                            r14h3_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_13":
                                        {
                                            r14h3_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_14":
                                        {
                                            r14h3_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_15":
                                        {
                                            r14h3_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h3_17":
                                        {
                                            r14h3_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_1":
                                        {
                                            r14h4_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_2":
                                        {
                                            r14h4_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_3":
                                        {
                                            r14h4_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_4":
                                        {
                                            r14h4_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_5":
                                        {
                                            r14h4_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_6":
                                        {
                                            r14h4_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_7":
                                        {
                                            r14h4_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_8":
                                        {
                                            r14h4_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_9":
                                        {
                                            r14h4_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_10":
                                        {
                                            r14h4_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_12":
                                        {
                                            r14h4_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_13":
                                        {
                                            r14h4_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_14":
                                        {
                                            r14h4_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_15":
                                        {
                                            r14h4_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r14h4_17":
                                        {
                                            r14h4_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }


                                }
                            }
                        }
                    }
                    if (UserID != null && Weight != 0)
                    {
                        iobj.insert_Dashboard_values(UserID, SurveyPeriod, MaritalStatus, Country, SurveyID, Gender, Weight, Location, AgeGroup, Education, Occupation, PersonalInc_BKK, HHIncome_BKK, PersonalInc_UPC, HHIncome_UPC, Period, S5_1, S5_2, S5_3, S5_4, S5_5, S5_6, S5_7, S5_8, S5_9, S5_10, S5_11, S5_12, S5_13, S5_14, S5_15, S5_16, S5_17, S5_18, S5_19, S6, S7_1, S7_2, S7_3, S7_4, S7_5, S7_6, S7_7, S7_8, S7_9, S7_10, S7_11, S7_12, S7_13, S7_14, S7_15, S7_16, S7_17, S7_18, S7_19, S8, S10_1, S11_1, BrTom, BrSpont_BBengRed, BrSpont_TivWafer, BrSpont_TivTwin, BrSpont_TivCombo, BrSpont_TivJumbo, BrSpont_Sanghai, BrSpont_BBengHazel, BrSpont_BBengMaxx, BrSpont_VoizWaffle, BrSpont_Voizwafflmini, BrSpont_Bissin, BrSpont_CalChees, BrAid_BBengRed, BrAid_TivWafer, BrAid_TivTwin, BrAid_TivCombo, BrAid_TivJumbo, BrAid_Sanghai, BrAid_BBengHazel, BrAid_BBengMaxx, BrAid_VoizWaffle, BrAid_Voizwafflmini, BrAid_Bissin, BrAid_CalChees, AdTom, AdSpont_BBengRed, AdSpont_TivWafer, AdSpont_TivTwin, AdSpont_TivCombo, AdSpont_TivJumbo, AdSpont_Sanghai, AdSpont_BBengHazel, AdSpont_BBengMaxx, AdSpont_VoizWaffle, AdSpont_Voizwafflmini, AdSpont_Bissin, AdSpont_CalChees, AdAid_BBengRed, AdAid_TivWafer, AdAid_TivTwin, AdAid_TivCombo, AdAid_TivJumbo, AdAid_Sanghai, AdAid_BBengHazel, AdAid_BBengMaxx, AdAid_VoizWaffle, AdAid_Voizwafflmini, AdAid_Bissin, AdAided_CalChees, EverCons_BBengRed, EverCons_TivWafer, EverCons_TivTwin, EverCons_TivCombo, EverCons_TivJumbo, EverCons_Sanghai, EverCons_BBengHazel, EverCons_BBengMaxx, EverCons_VoizWaffle, EverCons_Voizwafflmini, EverCons_Bissin, EverCons_CalChees, ConsL3M_BBengRed, ConsL3M_TivWafer, ConsL3M_TivTwin, ConsL3M_TivCombo, ConsL3M_TivJumbo, ConsL3M_Sanghai, ConsL3M_BBengHazel, ConsL3M_BBengMaxx, ConsL3M_VoizWaffle, ConsL3M_Voizwafflmini, ConsL3M_Bissin, ConsL3M_CalChees, ConsL1M_BBengRed, ConsL1M_TivWafer, ConsL1M_TivTwin, ConsL1M_TivCombo, ConsL1M_TivJumbo, ConsL1M_Sanghai, ConsL1M_BBengHazel, ConsL1M_BBengMaxx, ConsL1M_VoizWaffle, ConsL1M_Voizwafflmini, ConsL1M_Bissin, ConsL1M_CalChees, ConsL1W_BBengRed, ConsL1W_TivWafer, ConsL1W_TivTwin, ConsL1W_TivCombo, ConsL1W_TivJumbo, ConsL1W_Sanghai, ConsL1W_BBengHazel, ConsL1W_BBengMaxx, ConsL1W_VoizWaffle, ConsL1W_Voizwafflmini, ConsL1W_Bissin, ConsL1W_CalChees, Bumo, PreBumo, FavBrand_BBengRed, FavBrand_TivWafer, FavBrand_TivTwin, FavBrand_TivCombo, FavBrand_TivJumbo, FavBrand_Sanghai, FavBrand_BBengHazel, FavBrand_BBengMaxx, FavBrand_VoizWaffle, FavBrand_Voizwafflmini, FavBrand_Bissin, FavBrand_CalChees, MostFavBrand, NetBrTom, BrSpontNet_BBengNet, BrSpontNet_TivoliNet, BrSpontNet_SanghaiNet, BrSpontNet_VoizNet, BrSpontNet_BissinNet, BrAidNet_BBengNet, BrAidNet_TivoliNet, BrAidNet_SanghaiNet, BrAidNet_VoizNet, BrAidNet_BissinNet, NetAdTom, AdSpontNet_BBengNet, AdSpontNet_TivoliNet, AdSpontNet_SanghaiNet, AdSpontNet_VoizNet, AdSpontNet_BissinNet, AdAidNet_BBengNet, AdAidNet_TivoliNet, AdAidNet_SanghaiNet, AdAidNet_VoizNet, AdAidNet_BissinNet, EverConsNet_BBengNet, EverConsNet_TivoliNet, EverConsNet_SanghaiNet, EverConsNet_VoizNet, EverConsNet_BissinNet, ConsL3MNet_BBengNet, ConsL3MNet_TivoliNet, ConsL3MNet_SanghaiNet, ConsL3MNet_VoizNet, ConsL3MNet_BissinNet, ConsL1MNet_BBengNet, ConsL1MNet_TivoliNet, ConsL1MNet_SanghaiNet, ConsL1MNet_VoizNet, ConsL1MNet_BissinNet, ConsL1WNet_BBengNet, ConsL1WNet_TivoliNet, ConsL1WNet_SanghaiNet, ConsL1WNet_VoizNet, ConsL1WNet_BissinNet, Bumo_Net, PreBumo_Net, FavBrandNet_BBengNet, FavBrandNet_TivoliNet, FavBrandNet_SanghaiNet, FavBrandNet_VoizNet, FavBrandNet_BissinNet, MostFavBrand_Net, r14h1_1, r14h1_2, r14h1_3, r14h1_4, r14h1_5, r14h1_6, r14h1_7, r14h1_8, r14h1_9, r14h1_10, r14h1_12, r14h1_13, r14h1_14, r14h1_15, r14h1_17, r14h2_1, r14h2_2, r14h2_3, r14h2_4, r14h2_5, r14h2_6, r14h2_7, r14h2_8, r14h2_9, r14h2_10, r14h2_12, r14h2_13, r14h2_14, r14h2_15, r14h2_17, r14h3_1, r14h3_2, r14h3_3, r14h3_4, r14h3_5, r14h3_6, r14h3_7, r14h3_8, r14h3_9, r14h3_10, r14h3_12, r14h3_13, r14h3_14, r14h3_15, r14h3_17, r14h4_1, r14h4_2, r14h4_3, r14h4_4, r14h4_5, r14h4_6, r14h4_7, r14h4_8, r14h4_9, r14h4_10, r14h4_12, r14h4_13, r14h4_14, r14h4_15, r14h4_17);
                    }

                }

            }
        }

        private static string getYear(string SurveyPeriod)
        {
            string[] date = SurveyPeriod.Split('-');
            return date[0];
        }

        private static string find_UserID(int SurveyID, string SurveyPeriod, string u_id)
        {
            string sum = "";
            string[] date = SurveyPeriod.Split('-');
            foreach (string d in date)
            {
                sum = sum + d;

            }
            return u_id + SurveyID + sum;
        }

    }

 
 class ThaiInsertion 
    {
        SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
        internal void insert_Dashboard_variable_values(string VARIABLE_NAME, string VARIABLE_NAME_SUB_NAME, string VARIABLE_ID, string VARIABLE_VALUE, string VARIABLE_NAME_QUESTION, int SurveyID, string country, string BASE_VARIABLE_NAME, string SurveyPeriod)
        {
            connection.Open();
            SqlCommand cmd = new SqlCommand("INSERT INTO DashboardSPS_Variable_Values (VARIABLE_NAME,VARIABLE_NAME_SUB_NAME,VARIABLE_ID,VARIABLE_VALUE,VARIABLE_NAME_QUESTION,SURVEY_ID,SURVEY_COUNTRY,BASE_VARIABLE_NAME,SURVEY_PERIOD) " + "VALUES('" + VARIABLE_NAME + "','" + VARIABLE_NAME_SUB_NAME + "','" + VARIABLE_ID + "','" + VARIABLE_VALUE.Replace("'", "''") + "','" + VARIABLE_NAME_QUESTION + "','" + SurveyID + "','" + country + "','" + BASE_VARIABLE_NAME + "','" + SurveyPeriod + "')", connection);
            try
            {


                cmd.ExecuteNonQuery();
                Console.WriteLine("Dashboard variable values inserted successfully");

                connection.Close();



            }
            catch (Exception)
            {

                Console.WriteLine("Exception occured");
                Console.ReadLine();
            }

        }


        internal void insert_Dashboard_values(string UserID, string SurveyPeriod, string MaritalStatus, string Country, int SurveyID, string Gender, decimal Weight, string Location, string AgeGroup, string Education, string Occupation, string PersonalInc_BKK, string HHIncome_BKK, string PersonalInc_UPC, string HHIncome_UPC, string Period, string S5_1, string S5_2, string S5_3, string S5_4, string S5_5, string S5_6, string S5_7, string S5_8, string S5_9, string S5_10, string S5_11, string S5_12, string S5_13, string S5_14, string S5_15, string S5_16, string S5_17, string S5_18, string S5_19, string S6, string S7_1, string S7_2, string S7_3, string S7_4, string S7_5, string S7_6, string S7_7, string S7_8, string S7_9, string S7_10, string S7_11, string S7_12, string S7_13, string S7_14, string S7_15, string S7_16, string S7_17, string S7_18, string S7_19, string S8, string S10_1, string S11_1, string BrTom, string BrSpont_BBengRed, string BrSpont_TivWafer, string BrSpont_TivTwin, string BrSpont_TivCombo, string BrSpont_TivJumbo, string BrSpont_Sanghai, string BrSpont_BBengHazel, string BrSpont_BBengMaxx, string BrSpont_VoizWaffle, string BrSpont_Voizwafflmini, string BrSpont_Bissin, string BrSpont_CalChees, string BrAid_BBengRed, string BrAid_TivWafer, string BrAid_TivTwin, string BrAid_TivCombo, string BrAid_TivJumbo, string BrAid_Sanghai, string BrAid_BBengHazel, string BrAid_BBengMaxx, string BrAid_VoizWaffle, string BrAid_Voizwafflmini, string BrAid_Bissin, string BrAid_CalChees, string AdTom, string AdSpont_BBengRed, string AdSpont_TivWafer, string AdSpont_TivTwin, string AdSpont_TivCombo, string AdSpont_TivJumbo, string AdSpont_Sanghai, string AdSpont_BBengHazel, string AdSpont_BBengMaxx, string AdSpont_VoizWaffle, string AdSpont_Voizwafflmini, string AdSpont_Bissin, string AdSpont_CalChees, string AdAid_BBengRed, string AdAid_TivWafer, string AdAid_TivTwin, string AdAid_TivCombo, string AdAid_TivJumbo, string AdAid_Sanghai, string AdAid_BBengHazel, string AdAid_BBengMaxx, string AdAid_VoizWaffle, string AdAid_Voizwafflmini, string AdAid_Bissin, string AdAided_CalChees, string EverCons_BBengRed, string EverCons_TivWafer, string EverCons_TivTwin, string EverCons_TivCombo, string EverCons_TivJumbo, string EverCons_Sanghai, string EverCons_BBengHazel, string EverCons_BBengMaxx, string EverCons_VoizWaffle, string EverCons_Voizwafflmini, string EverCons_Bissin, string EverCons_CalChees, string ConsL3M_BBengRed, string ConsL3M_TivWafer, string ConsL3M_TivTwin, string ConsL3M_TivCombo, string ConsL3M_TivJumbo, string ConsL3M_Sanghai, string ConsL3M_BBengHazel, string ConsL3M_BBengMaxx, string ConsL3M_VoizWaffle, string ConsL3M_Voizwafflmini, string ConsL3M_Bissin, string ConsL3M_CalChees, string ConsL1M_BBengRed, string ConsL1M_TivWafer, string ConsL1M_TivTwin, string ConsL1M_TivCombo, string ConsL1M_TivJumbo, string ConsL1M_Sanghai, string ConsL1M_BBengHazel, string ConsL1M_BBengMaxx, string ConsL1M_VoizWaffle, string ConsL1M_Voizwafflmini, string ConsL1M_Bissin, string ConsL1M_CalChees, string ConsL1W_BBengRed, string ConsL1W_TivWafer, string ConsL1W_TivTwin, string ConsL1W_TivCombo, string ConsL1W_TivJumbo, string ConsL1W_Sanghai, string ConsL1W_BBengHazel, string ConsL1W_BBengMaxx, string ConsL1W_VoizWaffle, string ConsL1W_Voizwafflmini, string ConsL1W_Bissin, string ConsL1W_CalChees, string Bumo, string PreBumo, string FavBrand_BBengRed, string FavBrand_TivWafer, string FavBrand_TivTwin, string FavBrand_TivCombo, string FavBrand_TivJumbo, string FavBrand_Sanghai, string FavBrand_BBengHazel, string FavBrand_BBengMaxx, string FavBrand_VoizWaffle, string FavBrand_Voizwafflmini, string FavBrand_Bissin, string FavBrand_CalChees, string MostFavBrand, string NetBrTom, string BrSpontNet_BBengNet, string BrSpontNet_TivoliNet, string BrSpontNet_SanghaiNet, string BrSpontNet_VoizNet, string BrSpontNet_BissinNet, string BrAidNet_BBengNet, string BrAidNet_TivoliNet, string BrAidNet_SanghaiNet, string BrAidNet_VoizNet, string BrAidNet_BissinNet, string NetAdTom, string AdSpontNet_BBengNet, string AdSpontNet_TivoliNet, string AdSpontNet_SanghaiNet, string AdSpontNet_VoizNet, string AdSpontNet_BissinNet, string AdAidNet_BBengNet, string AdAidNet_TivoliNet, string AdAidNet_SanghaiNet, string AdAidNet_VoizNet, string AdAidNet_BissinNet, string EverConsNet_BBengNet, string EverConsNet_TivoliNet, string EverConsNet_SanghaiNet, string EverConsNet_VoizNet, string EverConsNet_BissinNet, string ConsL3MNet_BBengNet, string ConsL3MNet_TivoliNet, string ConsL3MNet_SanghaiNet, string ConsL3MNet_VoizNet, string ConsL3MNet_BissinNet, string ConsL1MNet_BBengNet, string ConsL1MNet_TivoliNet, string ConsL1MNet_SanghaiNet, string ConsL1MNet_VoizNet, string ConsL1MNet_BissinNet, string ConsL1WNet_BBengNet, string ConsL1WNet_TivoliNet, string ConsL1WNet_SanghaiNet, string ConsL1WNet_VoizNet, string ConsL1WNet_BissinNet, string Bumo_Net, string PreBumo_Net, string FavBrandNet_BBengNet, string FavBrandNet_TivoliNet, string FavBrandNet_SanghaiNet, string FavBrandNet_VoizNet, string FavBrandNet_BissinNet, string MostFavBrand_Net, string r14h1_1, string r14h1_2, string r14h1_3, string r14h1_4, string r14h1_5, string r14h1_6, string r14h1_7, string r14h1_8, string r14h1_9, string r14h1_10, string r14h1_12, string r14h1_13, string r14h1_14, string r14h1_15, string r14h1_17, string r14h2_1, string r14h2_2, string r14h2_3, string r14h2_4, string r14h2_5, string r14h2_6, string r14h2_7, string r14h2_8, string r14h2_9, string r14h2_10, string r14h2_12, string r14h2_13, string r14h2_14, string r14h2_15, string r14h2_17, string r14h3_1, string r14h3_2, string r14h3_3, string r14h3_4, string r14h3_5, string r14h3_6, string r14h3_7, string r14h3_8, string r14h3_9, string r14h3_10, string r14h3_12, string r14h3_13, string r14h3_14, string r14h3_15, string r14h3_17, string r14h4_1, string r14h4_2, string r14h4_3, string r14h4_4, string r14h4_5, string r14h4_6, string r14h4_7, string r14h4_8, string r14h4_9, string r14h4_10, string r14h4_12, string r14h4_13, string r14h4_14, string r14h4_15, string r14h4_17)
             
           {  
             int i;
             connection.Open();
             //SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
              SqlCommand cmd = new SqlCommand("INSERT INTO DashboardFLATT_THAIWAF_temp (UserID,AttendedOn,MaritalStatus,Country,SurveyID,Gender,Weight,Location,AgeGroup,Education,Occupation,PersonalInc_BKK,HHIncome_BKK,PersonalInc_UPC,HHIncome_UPC,Period,S5_1,S5_2,S5_3,S5_4,S5_5,S5_6,S5_7,S5_8,S5_9,S5_10,S5_11,S5_12,S5_13,S5_14,S5_15,S5_16,S5_17,S5_18,S5_19,S6,S7_1,S7_2,S7_3,S7_4,S7_5,S7_6,S7_7,S7_8,S7_9,S7_10,S7_11,S7_12,S7_13,S7_14,S7_15,S7_16,S7_17,S7_18,S7_19,S8,S10_1,S11_1,BrTom,BrSpont_BBengRed,BrSpont_TivWafer,BrSpont_TivTwin,BrSpont_TivCombo,BrSpont_TivJumbo,BrSpont_Sanghai,BrSpont_BBengHazel,BrSpont_BBengMaxx,BrSpont_VoizWaffle,BrSpont_Voizwafflmini,BrSpont_Bissin,BrSpont_CalChees,BrAid_BBengRed,BrAid_TivWafer,BrAid_TivTwin,BrAid_TivCombo,BrAid_TivJumbo,BrAid_Sanghai,BrAid_BBengHazel,BrAid_BBengMaxx,BrAid_VoizWaffle,BrAid_Voizwafflmini,BrAid_Bissin,BrAid_CalChees,AdTom,AdSpont_BBengRed,AdSpont_TivWafer,AdSpont_TivTwin,AdSpont_TivCombo,AdSpont_TivJumbo,AdSpont_Sanghai,AdSpont_BBengHazel,AdSpont_BBengMaxx,AdSpont_VoizWaffle,AdSpont_Voizwafflmini,AdSpont_Bissin,AdSpont_CalChees,AdAid_BBengRed,AdAid_TivWafer,AdAid_TivTwin,AdAid_TivCombo,AdAid_TivJumbo,AdAid_Sanghai,AdAid_BBengHazel,AdAid_BBengMaxx,AdAid_VoizWaffle,AdAid_Voizwafflmini,AdAid_Bissin,AdAided_CalChees,EverCons_BBengRed,EverCons_TivWafer,EverCons_TivTwin,EverCons_TivCombo,EverCons_TivJumbo,EverCons_Sanghai,EverCons_BBengHazel,EverCons_BBengMaxx,EverCons_VoizWaffle,EverCons_Voizwafflmini,EverCons_Bissin,EverCons_CalChees,ConsL3M_BBengRed,ConsL3M_TivWafer,ConsL3M_TivTwin,ConsL3M_TivCombo,ConsL3M_TivJumbo,ConsL3M_Sanghai,ConsL3M_BBengHazel,ConsL3M_BBengMaxx,ConsL3M_VoizWaffle,ConsL3M_Voizwafflmini,ConsL3M_Bissin,ConsL3M_CalChees,ConsL1M_BBengRed,ConsL1M_TivWafer,ConsL1M_TivTwin,ConsL1M_TivCombo,ConsL1M_TivJumbo,ConsL1M_Sanghai,ConsL1M_BBengHazel,ConsL1M_BBengMaxx,ConsL1M_VoizWaffle,ConsL1M_Voizwafflmini,ConsL1M_Bissin,ConsL1M_CalChees,ConsL1W_BBengRed,ConsL1W_TivWafer,ConsL1W_TivTwin,ConsL1W_TivCombo,ConsL1W_TivJumbo,ConsL1W_Sanghai,ConsL1W_BBengHazel,ConsL1W_BBengMaxx,ConsL1W_VoizWaffle,ConsL1W_Voizwafflmini,ConsL1W_Bissin,ConsL1W_CalChees,Bumo,PreBumo,FavBrand_BBengRed,FavBrand_TivWafer,FavBrand_TivTwin,FavBrand_TivCombo,FavBrand_TivJumbo,FavBrand_Sanghai,FavBrand_BBengHazel,FavBrand_BBengMaxx,FavBrand_VoizWaffle,FavBrand_Voizwafflmini,FavBrand_Bissin,FavBrand_CalChees,MostFavBrand,NetBrTom,BrSpontNet_BBengNet,BrSpontNet_TivoliNet,BrSpontNet_SanghaiNet,BrSpontNet_VoizNet,BrSpontNet_BissinNet,BrAidNet_BBengNet,BrAidNet_TivoliNet,BrAidNet_SanghaiNet,BrAidNet_VoizNet,BrAidNet_BissinNet,NetAdTom,AdSpontNet_BBengNet,AdSpontNet_TivoliNet,AdSpontNet_SanghaiNet,AdSpontNet_VoizNet,AdSpontNet_BissinNet,AdAidNet_BBengNet,AdAidNet_TivoliNet,AdAidNet_SanghaiNet,AdAidNet_VoizNet,AdAidNet_BissinNet,EverConsNet_BBengNet,EverConsNet_TivoliNet,EverConsNet_SanghaiNet,EverConsNet_VoizNet,EverConsNet_BissinNet,ConsL3MNet_BBengNet,ConsL3MNet_TivoliNet,ConsL3MNet_SanghaiNet,ConsL3MNet_VoizNet,ConsL3MNet_BissinNet,ConsL1MNet_BBengNet,ConsL1MNet_TivoliNet,ConsL1MNet_SanghaiNet,ConsL1MNet_VoizNet,ConsL1MNet_BissinNet,ConsL1WNet_BBengNet,ConsL1WNet_TivoliNet,ConsL1WNet_SanghaiNet,ConsL1WNet_VoizNet,ConsL1WNet_BissinNet,Bumo_Net,PreBumo_Net,FavBrandNet_BBengNet,FavBrandNet_TivoliNet,FavBrandNet_SanghaiNet,FavBrandNet_VoizNet,FavBrandNet_BissinNet,MostFavBrand_Net,r14h1_1,r14h1_2,r14h1_3,r14h1_4,r14h1_5,r14h1_6,r14h1_7,r14h1_8,r14h1_9,r14h1_10,r14h1_12,r14h1_13,r14h1_14,r14h1_15,r14h1_17,r14h2_1,r14h2_2,r14h2_3,r14h2_4,r14h2_5,r14h2_6,r14h2_7,r14h2_8,r14h2_9,r14h2_10,r14h2_12,r14h2_13,r14h2_14,r14h2_15,r14h2_17,r14h3_1,r14h3_2,r14h3_3,r14h3_4,r14h3_5,r14h3_6,r14h3_7,r14h3_8,r14h3_9,r14h3_10,r14h3_12,r14h3_13,r14h3_14,r14h3_15,r14h3_17,r14h4_1,r14h4_2,r14h4_3,r14h4_4,r14h4_5,r14h4_6,r14h4_7,r14h4_8,r14h4_9,r14h4_10,r14h4_12,r14h4_13,r14h4_14,r14h4_15,r14h4_17) " + "VALUES('" + UserID + "','" + SurveyPeriod + "','" + MaritalStatus + "','" + Country + "','" + SurveyID + "','" + Gender + "','" + Weight + "','" + Location + "','" + AgeGroup + "','" + Education + "','" + Occupation + "','" + PersonalInc_BKK + "','" + HHIncome_BKK + "','" + PersonalInc_UPC + "','" + HHIncome_UPC + "','" + Period + "','" + S5_1 + "','" + S5_2 + "','" + S5_3 + "','" + S5_4 + "','" + S5_5 + "','" + S5_6 + "','" + S5_7 + "','" + S5_8 + "','" + S5_9 + "','" + S5_10 + "','" + S5_11 + "','" + S5_12 + "','" + S5_13 + "','" + S5_14 + "','" + S5_15 + "','" + S5_16 + "','" + S5_17 + "','" + S5_18 + "','" + S5_19 + "','" + S6 + "','" + S7_1 + "','" + S7_2 + "','" + S7_3 + "','" + S7_4 + "','" + S7_5 + "','" + S7_6 + "','" + S7_7 + "','" + S7_8 + "','" + S7_9 + "','" + S7_10 + "','" + S7_11 + "','" + S7_12 + "','" + S7_13 + "','" + S7_14 + "','" + S7_15 + "','" + S7_16 + "','" + S7_17 + "','" + S7_18 + "','" + S7_19 + "','" + S8 + "','" + S10_1 + "','" + S11_1 + "','" + BrTom + "','" + BrSpont_BBengRed + "','" + BrSpont_TivWafer + "','" + BrSpont_TivTwin + "','" + BrSpont_TivCombo + "','" + BrSpont_TivJumbo + "','" + BrSpont_Sanghai + "','" + BrSpont_BBengHazel + "','" + BrSpont_BBengMaxx + "','" + BrSpont_VoizWaffle + "','" + BrSpont_Voizwafflmini + "','" + BrSpont_Bissin + "','" + BrSpont_CalChees + "','" + BrAid_BBengRed + "','" + BrAid_TivWafer + "','" + BrAid_TivTwin + "','" + BrAid_TivCombo + "','" + BrAid_TivJumbo + "','" + BrAid_Sanghai + "','" + BrAid_BBengHazel + "','" + BrAid_BBengMaxx + "','" + BrAid_VoizWaffle + "','" + BrAid_Voizwafflmini + "','" + BrAid_Bissin + "','" + BrAid_CalChees + "','" + AdTom + "','" + AdSpont_BBengRed + "','" + AdSpont_TivWafer + "','" + AdSpont_TivTwin + "','" + AdSpont_TivCombo + "','" + AdSpont_TivJumbo + "','" + AdSpont_Sanghai + "','" + AdSpont_BBengHazel + "','" + AdSpont_BBengMaxx + "','" + AdSpont_VoizWaffle + "','" + AdSpont_Voizwafflmini + "','" + AdSpont_Bissin + "','" + AdSpont_CalChees + "','" + AdAid_BBengRed + "','" + AdAid_TivWafer + "','" + AdAid_TivTwin + "','" + AdAid_TivCombo + "','" + AdAid_TivJumbo + "','" + AdAid_Sanghai + "','" + AdAid_BBengHazel + "','" + AdAid_BBengMaxx + "','" + AdAid_VoizWaffle + "','" + AdAid_Voizwafflmini + "','" + AdAid_Bissin + "','" + AdAided_CalChees + "','" + EverCons_BBengRed + "','" + EverCons_TivWafer + "','" + EverCons_TivTwin + "','" + EverCons_TivCombo + "','" + EverCons_TivJumbo + "','" + EverCons_Sanghai + "','" + EverCons_BBengHazel + "','" + EverCons_BBengMaxx + "','" + EverCons_VoizWaffle + "','" + EverCons_Voizwafflmini + "','" + EverCons_Bissin + "','" + EverCons_CalChees + "','" + ConsL3M_BBengRed + "','" + ConsL3M_TivWafer + "','" + ConsL3M_TivTwin + "','" + ConsL3M_TivCombo + "','" + ConsL3M_TivJumbo + "','" + ConsL3M_Sanghai + "','" + ConsL3M_BBengHazel + "','" + ConsL3M_BBengMaxx + "','" + ConsL3M_VoizWaffle + "','" + ConsL3M_Voizwafflmini + "','" + ConsL3M_Bissin + "','" + ConsL3M_CalChees + "','" + ConsL1M_BBengRed + "','" + ConsL1M_TivWafer + "','" + ConsL1M_TivTwin + "','" + ConsL1M_TivCombo + "','" + ConsL1M_TivJumbo + "','" + ConsL1M_Sanghai + "','" + ConsL1M_BBengHazel + "','" + ConsL1M_BBengMaxx + "','" + ConsL1M_VoizWaffle + "','" + ConsL1M_Voizwafflmini + "','" + ConsL1M_Bissin + "','" + ConsL1M_CalChees + "','" + ConsL1W_BBengRed + "','" + ConsL1W_TivWafer + "','" + ConsL1W_TivTwin + "','" + ConsL1W_TivCombo + "','" + ConsL1W_TivJumbo + "','" + ConsL1W_Sanghai + "','" + ConsL1W_BBengHazel + "','" + ConsL1W_BBengMaxx + "','" + ConsL1W_VoizWaffle + "','" + ConsL1W_Voizwafflmini + "','" + ConsL1W_Bissin + "','" + ConsL1W_CalChees + "','" + Bumo + "','" + PreBumo + "','" + FavBrand_BBengRed + "','" + FavBrand_TivWafer + "','" + FavBrand_TivTwin + "','" + FavBrand_TivCombo + "','" + FavBrand_TivJumbo + "','" + FavBrand_Sanghai + "','" + FavBrand_BBengHazel + "','" + FavBrand_BBengMaxx + "','" + FavBrand_VoizWaffle + "','" + FavBrand_Voizwafflmini + "','" + FavBrand_Bissin + "','" + FavBrand_CalChees + "','" + MostFavBrand + "','" + NetBrTom + "','" + BrSpontNet_BBengNet + "','" + BrSpontNet_TivoliNet + "','" + BrSpontNet_SanghaiNet + "','" + BrSpontNet_VoizNet + "','" + BrSpontNet_BissinNet + "','" + BrAidNet_BBengNet + "','" + BrAidNet_TivoliNet + "','" + BrAidNet_SanghaiNet + "','" + BrAidNet_VoizNet + "','" + BrAidNet_BissinNet + "','" + NetAdTom + "','" + AdSpontNet_BBengNet + "','" + AdSpontNet_TivoliNet + "','" + AdSpontNet_SanghaiNet + "','" + AdSpontNet_VoizNet + "','" + AdSpontNet_BissinNet + "','" + AdAidNet_BBengNet + "','" + AdAidNet_TivoliNet + "','" + AdAidNet_SanghaiNet + "','" + AdAidNet_VoizNet + "','" + AdAidNet_BissinNet + "','" + EverConsNet_BBengNet + "','" + EverConsNet_TivoliNet + "','" + EverConsNet_SanghaiNet + "','" + EverConsNet_VoizNet + "','" + EverConsNet_BissinNet + "','" + ConsL3MNet_BBengNet + "','" + ConsL3MNet_TivoliNet + "','" + ConsL3MNet_SanghaiNet + "','" + ConsL3MNet_VoizNet + "','" + ConsL3MNet_BissinNet + "','" + ConsL1MNet_BBengNet + "','" + ConsL1MNet_TivoliNet + "','" + ConsL1MNet_SanghaiNet + "','" + ConsL1MNet_VoizNet + "','" + ConsL1MNet_BissinNet + "','" + ConsL1WNet_BBengNet + "','" + ConsL1WNet_TivoliNet + "','" + ConsL1WNet_SanghaiNet + "','" + ConsL1WNet_VoizNet + "','" + ConsL1WNet_BissinNet + "','" + Bumo_Net + "','" + PreBumo_Net + "','" + FavBrandNet_BBengNet + "','" + FavBrandNet_TivoliNet + "','" + FavBrandNet_SanghaiNet + "','" + FavBrandNet_VoizNet + "','" + FavBrandNet_BissinNet + "','" + MostFavBrand_Net + "','" + r14h1_1 + "','" + r14h1_2 + "','" + r14h1_3 + "','" + r14h1_4 + "','" + r14h1_5 + "','" + r14h1_6 + "','" + r14h1_7 + "','" + r14h1_8 + "','" + r14h1_9 + "','" + r14h1_10 + "','" + r14h1_12 + "','" + r14h1_13 + "','" + r14h1_14 + "','" + r14h1_15 + "','" + r14h1_17 + "','" + r14h2_1 + "','" + r14h2_2 + "','" + r14h2_3 + "','" + r14h2_4 + "','" + r14h2_5 + "','" + r14h2_6 + "','" + r14h2_7 + "','" + r14h2_8 + "','" + r14h2_9 + "','" + r14h2_10 + "','" + r14h2_12 + "','" + r14h2_13 + "','" + r14h2_14 + "','" + r14h2_15 + "','" + r14h2_17 + "','" + r14h3_1 + "','" + r14h3_2 + "','" + r14h3_3 + "','" + r14h3_4 + "','" + r14h3_5 + "','" + r14h3_6 + "','" + r14h3_7 + "','" + r14h3_8 + "','" + r14h3_9 + "','" + r14h3_10 + "','" + r14h3_12 + "','" + r14h3_13 + "','" + r14h3_14 + "','" + r14h3_15 + "','" + r14h3_17 + "','" + r14h4_1 + "','" + r14h4_2 + "','" + r14h4_3 + "','" + r14h4_4 + "','" + r14h4_5 + "','" + r14h4_6 + "','" + r14h4_7 + "','" + r14h4_8 + "','" + r14h4_9 + "','" + r14h4_10 + "','" + r14h4_12 + "','" + r14h4_13 + "','" + r14h4_14 + "','" + r14h4_15 + "','" + r14h4_17 + "' )", connection);
           try
            {
                cmd.ExecuteNonQuery();
                Console.WriteLine("Data inserted successfully");
                i = 1;

            }
            catch (Exception ex)
            {

                connection.Close();
                i = 0;
                Console.WriteLine("Exception occured" + ex.ToString());
                Console.WriteLine("Exception occured");

                Console.ReadLine();
            }
            connection.Close();
        }


    }

}
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZambiaDataManager.modules.mcsp_lng
{
    public class discontinueIudIndicatorRow : ISerialisableWebDataset
    {
        //we get the data for each indicator
        [JsonIgnore] public string sys_entered_by;
        [JsonIgnore] public string sys_data_entry_date;
        [JsonProperty("sys-data-edit-date")] public string sys_data_edit_date;
        [JsonProperty("record-serial-number")] public string record_serial_number;
        [JsonProperty("province-name")] public string province_name;
        [JsonProperty("district-name")] public string district_name;
        [JsonProperty("facility-name")] public string facility_name;
        [JsonProperty("client-name")] public string client_name;
        [JsonProperty("year-of-birth")] public string year_of_birth;
        [JsonProperty("family-planning-card-no")] public string family_planning_card_no;
        [JsonProperty("date-of-service")] public string date_of_service;
        [JsonProperty("marital-status")] public string marital_status;
        [JsonProperty("other-marital-status-specified")] public string other_marital_status_specified;
        [JsonProperty("number-pregnancies")] public string number_pregnancies;
        [JsonProperty("education-level")] public string education_level;
        [JsonProperty("type-of-iud-received")] public string type_of_iud_received;
        [JsonProperty("length-iud-use")] public string length_iud_use;
        [JsonIgnore] public string ask_ques_participant;
        [JsonIgnore] public string label_why_decided_iud_removed;
        [JsonProperty("iud-removed-not-specified")] public string iud_removed_not_specified;
        [JsonProperty("iud-removed-iam-pregnant")] public string iud_removed_iam_pregnant;
        [JsonProperty("iud-removed-want-to-get-pregnant")] public string iud_removed_want_to_get_pregnant;
        [JsonProperty("iud-removed-decreased-menstrual-amenorrhea")] public string iud_removed_decreased_menstrual_amenorrhea;
        [JsonProperty("iud-removed-bleeding-disturbances")] public string iud_removed_bleeding_disturbances;
        [JsonProperty("iud-removed-pain-with-method")] public string iud_removed_pain_with_method;
        [JsonProperty("iud-removed-weight-gain")] public string iud_removed_weight_gain;
        [JsonProperty("iud-removed-afraid-infertile")] public string iud_removed_afraid_infertile;
        [JsonProperty("iud-removed-husband-didnt-want")] public string iud_removed_husband_didnt_want;
        [JsonProperty("iud-removed-family-mbrs-didnt-want")] public string iud_removed_family_mbrs_didnt_want;
        [JsonProperty("iud-removed-not-sexually-active")] public string iud_removed_not_sexually_active;
        [JsonProperty("iud-removed-menopausal")] public string iud_removed_menopausal;
        [JsonProperty("iud-removed-other")] public string iud_removed_other;
        [JsonProperty("iud-removed-other-specified")] public string iud_removed_other_specified;
        [JsonIgnore] public string label_family_planning_to_use;
        [JsonProperty("fp-pref-not-specified")] public string fp_pref_not_specified;
        [JsonProperty("fp-pref-no-method")] public string fp_pref_no_method;
        [JsonProperty("fp-pref-preg-trying-toget")] public string fp_pref_preg_trying_toget;
        [JsonProperty("fp-pref-pills")] public string fp_pref_pills;
        [JsonProperty("fp-pref-implant")] public string fp_pref_implant;
        [JsonProperty("fp-pref-injectables")] public string fp_pref_injectables;
        [JsonProperty("fp-pref-condoms-only")] public string fp_pref_condoms_only;
        [JsonProperty("fp-pref-copper-iud")] public string fp_pref_copper_iud;
        [JsonProperty("fp-pref-hormnal-iud")] public string fp_pref_hormnal_iud;
        [JsonProperty("fp-pref-female-sterilisation")] public string fp_pref_female_sterilisation;
        [JsonProperty("fp-pref-male-sterilisation")] public string fp_pref_male_sterilisation;
        [JsonProperty("fp-pref-emergency-contraception")] public string fp_pref_emergency_contraception;
        [JsonProperty("fp-pref-cycle-beads")] public string fp_pref_cycle_beads;
        [JsonProperty("fp-pref-traditional-method")] public string fp_pref_traditional_method;


        public DataTable getTable()
        {
            var table = new DataTable();
            var fields = new List<string>{
                //"sys-entered-by",
//"sys-data-entry-date",
"sys-data-edit-date",
"record-serial-number",
"province-name",
"district-name",
"facility-name",
"client-name",
"year-of-birth",
"family-planning-card-no",
"date-of-service",
"marital-status",
"other-marital-status-specified",
"number-pregnancies",
"education-level",
"type-of-iud-received",
"length-iud-use",
//"ask-ques-participant",
//"label-why-decided-iud-removed",
"iud-removed-not-specified",
"iud-removed-iam-pregnant",
"iud-removed-want-to-get-pregnant",
"iud-removed-decreased-menstrual-amenorrhea",
"iud-removed-bleeding-disturbances",
"iud-removed-pain-with-method",
"iud-removed-weight-gain",
"iud-removed-afraid-infertile",
"iud-removed-husband-didnt-want",
"iud-removed-family-mbrs-didnt-want",
"iud-removed-not-sexually-active",
"iud-removed-menopausal",
"iud-removed-other",
"iud-removed-other-specified",
//"label-family-planning-to-use",
"fp-pref-not-specified",
"fp-pref-no-method",
"fp-pref-preg-trying-toget",
"fp-pref-pills",
"fp-pref-implant",
"fp-pref-injectables",
"fp-pref-condoms-only",
"fp-pref-copper-iud",
"fp-pref-hormnal-iud",
"fp-pref-female-sterilisation",
"fp-pref-male-sterilisation",
"fp-pref-emergency-contraception",
"fp-pref-cycle-beads",
"fp-pref-traditional-method"
};
            fields.ForEach(t => table.Columns.Add(t.Replace('-', '_')));
            return table;
        }

        public DataRow toRow(DataRow row)
        {
            //row["sys_entered_by"] = sys_entered_by;
            //row["sys_data_entry_date"] = sys_data_entry_date;
            row["sys_data_edit_date"] = sys_data_edit_date;
            row["record_serial_number"] = record_serial_number;
            row["province_name"] = province_name;
            row["district_name"] = district_name;
            row["facility_name"] = facility_name;
            row["client_name"] = client_name;
            row["year_of_birth"] = year_of_birth;
            row["family_planning_card_no"] = family_planning_card_no;
            row["date_of_service"] = date_of_service;
            row["marital_status"] = marital_status;
            row["other_marital_status_specified"] = other_marital_status_specified;
            row["number_pregnancies"] = number_pregnancies;
            row["education_level"] = education_level;
            row["type_of_iud_received"] = type_of_iud_received;
            row["length_iud_use"] = length_iud_use;
            //row["ask_ques_participant"] = ask_ques_participant;
            //row["label_why_decided_iud_removed"] = label_why_decided_iud_removed;
            row["iud_removed_not_specified"] = iud_removed_not_specified;
            row["iud_removed_iam_pregnant"] = iud_removed_iam_pregnant;
            row["iud_removed_want_to_get_pregnant"] = iud_removed_want_to_get_pregnant;
            row["iud_removed_decreased_menstrual_amenorrhea"] = iud_removed_decreased_menstrual_amenorrhea;
            row["iud_removed_bleeding_disturbances"] = iud_removed_bleeding_disturbances;
            row["iud_removed_pain_with_method"] = iud_removed_pain_with_method;
            row["iud_removed_weight_gain"] = iud_removed_weight_gain;
            row["iud_removed_afraid_infertile"] = iud_removed_afraid_infertile;
            row["iud_removed_husband_didnt_want"] = iud_removed_husband_didnt_want;
            row["iud_removed_family_mbrs_didnt_want"] = iud_removed_family_mbrs_didnt_want;
            row["iud_removed_not_sexually_active"] = iud_removed_not_sexually_active;
            row["iud_removed_menopausal"] = iud_removed_menopausal;
            row["iud_removed_other"] = iud_removed_other;
            row["iud_removed_other_specified"] = iud_removed_other_specified;
            //row["label_family_planning_to_use"] = label_family_planning_to_use;
            row["fp_pref_not_specified"] = fp_pref_not_specified;
            row["fp_pref_no_method"] = fp_pref_no_method;
            row["fp_pref_preg_trying_toget"] = fp_pref_preg_trying_toget;
            row["fp_pref_pills"] = fp_pref_pills;
            row["fp_pref_implant"] = fp_pref_implant;
            row["fp_pref_injectables"] = fp_pref_injectables;
            row["fp_pref_condoms_only"] = fp_pref_condoms_only;
            row["fp_pref_copper_iud"] = fp_pref_copper_iud;
            row["fp_pref_hormnal_iud"] = fp_pref_hormnal_iud;
            row["fp_pref_female_sterilisation"] = fp_pref_female_sterilisation;
            row["fp_pref_male_sterilisation"] = fp_pref_male_sterilisation;
            row["fp_pref_emergency_contraception"] = fp_pref_emergency_contraception;
            row["fp_pref_cycle_beads"] = fp_pref_cycle_beads;
            row["fp_pref_traditional_method"] = fp_pref_traditional_method;

            return row;
        }
    }
}

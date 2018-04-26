using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZambiaDataManager.modules.mcsp_lng
{
    public class receivingIudIndicatorRow : ISerialisableWebDataset
    {
        //we get the data for each indicator
        [JsonProperty("sys-entered-by")] public string sys_entered_by;
        [JsonProperty("sys-data-entry-date")] public string sys_data_entry_date;
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
        [JsonProperty("number-births")] public string number_births;
        [JsonProperty("education-level")] public string education_level;
        [JsonProperty("type-of-iud-received")] public string type_of_iud_received;
        [JsonProperty("visit-type")] public string visit_type;
        [JsonProperty("method-used-until-today")] public string method_used_until_today;
        [JsonProperty("other-method-specified")] public string other_method_specified;
        [JsonProperty("timing")] public string timing;
        [JsonProperty("label-iud-reasons")] public string label_iud_reasons;
        [JsonProperty("iud-reas-not-specified")] public string iud_reas_not_specified;
        [JsonProperty("iud-reas-reduces-menstrual-bleeding")] public string iud_reas_reduces_menstrual_bleeding;
        [JsonProperty("iud-reas-reversible")] public string iud_reas_reversible;
        [JsonProperty("iud-reas-convinient")] public string iud_reas_convinient;
        [JsonProperty("iud-reas-delay-pregnancy")] public string iud_reas_delay_pregnancy;
        [JsonProperty("iud-reas-dont-want-more-children")] public string iud_reas_dont_want_more_children;
        [JsonProperty("iud-reas-fewer-side-effects")] public string iud_reas_fewer_side_effects;
        [JsonProperty("iud-reas-discreet")] public string iud_reas_discreet;
        [JsonProperty("iud-reas-ok-use-while-breastfeeding")] public string iud_reas_ok_use_while_breastfeeding;
        [JsonProperty("iud-reas-more-affordable")] public string iud_reas_more_affordable;
        [JsonProperty("iud-reas-lasts-longer")] public string iud_reas_lasts_longer;
        [JsonProperty("iud-reas-highly-effective")] public string iud_reas_highly_effective;
        [JsonProperty("iud-reas-friend-family-recommended")] public string iud_reas_friend_family_recommended;
        [JsonProperty("iud-reas-not-sure")] public string iud_reas_not_sure;
        [JsonProperty("iud-reas-other")] public string iud_reas_other;
        [JsonProperty("iud-reas-other-specified")] public string iud_reas_other_specified;
        [JsonProperty("label-fp-tohave-chosen")] public string label_fp_tohave_chosen;
        [JsonProperty("fp-pref-not-specified")] public string fp_pref_not_specified;
        [JsonProperty("fp-pref-no-method")] public string fp_pref_no_method;
        [JsonProperty("fp-pref-implant")] public string fp_pref_implant;
        [JsonProperty("fp-pref-injectables")] public string fp_pref_injectables;
        [JsonProperty("fp-pref-other-iud")] public string fp_pref_other_iud;
        [JsonProperty("fp-pref-pills")] public string fp_pref_pills;
        [JsonProperty("fp-pref-condoms-only")] public string fp_pref_condoms_only;
        [JsonProperty("fp-pref-emergency-contraception")] public string fp_pref_emergency_contraception;
        [JsonProperty("fp-pref-traditional-method")] public string fp_pref_traditional_method;
        [JsonProperty("fp-pref-lam")] public string fp_pref_lam;
        [JsonProperty("fp-pref-other")] public string fp_pref_other;
        [JsonProperty("fp-pref-other-specified")] public string fp_pref_other_specified;
        [JsonProperty("heard-of-hormonal-iud")] public string heard_of_hormonal_iud;
        [JsonProperty("label-first-found-out")] public string label_first_found_out;
        [JsonProperty("first-found-out-not-specified")] public string first_found_out_not_specified;
        [JsonProperty("first-found-out-health-care-worker")] public string first_found_out_health_care_worker;
        [JsonProperty("first-found-out-community-health-worker")] public string first_found_out_community_health_worker;
        [JsonProperty("first-found-out-friend-or-family")] public string first_found_out_friend_or_family;
        [JsonProperty("first-found-out-radio-or-tv")] public string first_found_out_radio_or_tv;
        [JsonProperty("first-found-out-poster-or-flyer")] public string first_found_out_poster_or_flyer;
        [JsonProperty("first-found-out-other")] public string first_found_out_other;
        [JsonProperty("first-found-out-other-specified")] public string first_found_out_other_specified;
        [JsonProperty("label-are-you-willing")] public string label_are_you_willing;
        [JsonProperty("refused-unable-to-give-phone")] public string refused_unable_to_give_phone;
        [JsonProperty("label-if-willing")] public string label_if_willing;
        [JsonProperty("telephone-number")] public string telephone_number;
        [JsonProperty("preferred-language")] public string preferred_language;


        public DataTable getTable()
        {
            var table = new DataTable();
            var fields = new List<string>{"sys-entered-by",
"sys-data-entry-date",
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
"number-births",
"education-level",
"type-of-iud-received",
"visit-type",
"method-used-until-today",
"other-method-specified",
"timing",
"label-iud-reasons",
"iud-reas-not-specified",
"iud-reas-reduces-menstrual-bleeding",
"iud-reas-reversible",
"iud-reas-convinient",
"iud-reas-delay-pregnancy",
"iud-reas-dont-want-more-children",
"iud-reas-fewer-side-effects",
"iud-reas-discreet",
"iud-reas-ok-use-while-breastfeeding",
"iud-reas-more-affordable",
"iud-reas-lasts-longer",
"iud-reas-highly-effective",
"iud-reas-friend-family-recommended",
"iud-reas-not-sure",
"iud-reas-other",
"iud-reas-other-specified",
"label-fp-tohave-chosen",
"fp-pref-not-specified",
"fp-pref-no-method",
"fp-pref-implant",
"fp-pref-injectables",
"fp-pref-other-iud",
"fp-pref-pills",
"fp-pref-condoms-only",
"fp-pref-emergency-contraception",
"fp-pref-traditional-method",
"fp-pref-lam",
"fp-pref-other",
"fp-pref-other-specified",
"heard-of-hormonal-iud",
"label-first-found-out",
"first-found-out-not-specified",
"first-found-out-health-care-worker",
"first-found-out-community-health-worker",
"first-found-out-friend-or-family",
"first-found-out-radio-or-tv",
"first-found-out-poster-or-flyer",
"first-found-out-other",
"first-found-out-other-specified",
"label-are-you-willing",
"refused-unable-to-give-phone",
"label-if-willing",
"telephone-number",
"preferred-language",
};
            fields.ForEach(t => table.Columns.Add(t.Replace('-', '_')));
            return table;
        }

        public DataRow toRow(DataRow row)
        {
            row["sys_entered_by"] = sys_entered_by;
            row["sys_data_entry_date"] = sys_data_entry_date;
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
            row["number_births"] = number_births;
            row["education_level"] = education_level;
            row["type_of_iud_received"] = type_of_iud_received;
            row["visit_type"] = visit_type;
            row["method_used_until_today"] = method_used_until_today;
            row["other_method_specified"] = other_method_specified;
            row["timing"] = timing;
            row["label_iud_reasons"] = label_iud_reasons;
            row["iud_reas_not_specified"] = iud_reas_not_specified;
            row["iud_reas_reduces_menstrual_bleeding"] = iud_reas_reduces_menstrual_bleeding;
            row["iud_reas_reversible"] = iud_reas_reversible;
            row["iud_reas_convinient"] = iud_reas_convinient;
            row["iud_reas_delay_pregnancy"] = iud_reas_delay_pregnancy;
            row["iud_reas_dont_want_more_children"] = iud_reas_dont_want_more_children;
            row["iud_reas_fewer_side_effects"] = iud_reas_fewer_side_effects;
            row["iud_reas_discreet"] = iud_reas_discreet;
            row["iud_reas_ok_use_while_breastfeeding"] = iud_reas_ok_use_while_breastfeeding;
            row["iud_reas_more_affordable"] = iud_reas_more_affordable;
            row["iud_reas_lasts_longer"] = iud_reas_lasts_longer;
            row["iud_reas_highly_effective"] = iud_reas_highly_effective;
            row["iud_reas_friend_family_recommended"] = iud_reas_friend_family_recommended;
            row["iud_reas_not_sure"] = iud_reas_not_sure;
            row["iud_reas_other"] = iud_reas_other;
            row["iud_reas_other_specified"] = iud_reas_other_specified;
            row["label_fp_tohave_chosen"] = label_fp_tohave_chosen;
            row["fp_pref_not_specified"] = fp_pref_not_specified;
            row["fp_pref_no_method"] = fp_pref_no_method;
            row["fp_pref_implant"] = fp_pref_implant;
            row["fp_pref_injectables"] = fp_pref_injectables;
            row["fp_pref_other_iud"] = fp_pref_other_iud;
            row["fp_pref_pills"] = fp_pref_pills;
            row["fp_pref_condoms_only"] = fp_pref_condoms_only;
            row["fp_pref_emergency_contraception"] = fp_pref_emergency_contraception;
            row["fp_pref_traditional_method"] = fp_pref_traditional_method;
            row["fp_pref_lam"] = fp_pref_lam;
            row["fp_pref_other"] = fp_pref_other;
            row["fp_pref_other_specified"] = fp_pref_other_specified;
            row["heard_of_hormonal_iud"] = heard_of_hormonal_iud;
            row["label_first_found_out"] = label_first_found_out;
            row["first_found_out_not_specified"] = first_found_out_not_specified;
            row["first_found_out_health_care_worker"] = first_found_out_health_care_worker;
            row["first_found_out_community_health_worker"] = first_found_out_community_health_worker;
            row["first_found_out_friend_or_family"] = first_found_out_friend_or_family;
            row["first_found_out_radio_or_tv"] = first_found_out_radio_or_tv;
            row["first_found_out_poster_or_flyer"] = first_found_out_poster_or_flyer;
            row["first_found_out_other"] = first_found_out_other;
            row["first_found_out_other_specified"] = first_found_out_other_specified;
            row["label_are_you_willing"] = label_are_you_willing;
            row["refused_unable_to_give_phone"] = refused_unable_to_give_phone;
            row["label_if_willing"] = label_if_willing;
            row["telephone_number"] = telephone_number;
            row["preferred_language"] = preferred_language;

            return row;
        }
    }
}

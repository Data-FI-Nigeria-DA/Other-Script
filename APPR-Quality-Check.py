import pandas as pd

def validate_hts_prep_data_separate(excel_file_path, output_file_path):
    """
    Validates data in an Excel file, adds a "Validation Result" column,
    and saves the modified data to a new Excel file.
    Validation results are specific to each indicator category.

    Args:
        excel_file_path (str): The path to the input Excel file.
        output_file_path (str): The path to the output Excel file.
    """
    try:
        df = pd.read_excel(excel_file_path)

        # Define the indicator categories
        hts_self_user_categories = [
            "Qtr HTS_SELF (HIVSelfTestUser) Unassisted, Caregiver for Child",
            "Qtr HTS_SELF (HIVSelfTestUser) Unassisted, Other",
            "Qtr HTS_SELF (HIVSelfTestUser) Unassisted, Self",
            "Qtr HTS_SELF (HIVSelfTestUser) Unassisted, Sex Partner",
        ]

        hts_self_categories = [
            "Qtr_HTS_SELF (HIVSelfTest) 10-14yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 10-14yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 15-19yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 15-19yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 20-24yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 20-24yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 25-29yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 25-29yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 30-34yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 30-34yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 35-39yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 35-39yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 40-44yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 40-44yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 45-49yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 45-49yrs, Male, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 50+yrs, Female, Unassisted",
            "Qtr_HTS_SELF (HIVSelfTest) 50+yrs, Male, Unassisted",
        ]


        # Define the indicator categories
        PrEP_CT_PregBF_categories= [
            "Qtr PrEP_CT (PregBF ) Female, Pregnant",
            "Qtr PrEP_CT (PregBF ) Female,  Breastfeeding",
        ]

        PrEP_CT_Distribution_categories= [
            "Qtr PrEP_CT. (Distribution) Facility",
            "Qtr PrEP_CT. (Distribution) Community",
        ]

        PrEP_CT_TestResult_categories= [
            "Qtr PrEP_CT. (TestResult) Positive",
            "Qtr PrEP_CT. (TestResult) Other Test Result",
            "Qtr PrEP_CT. (TestResult) Negative"
        ]

        PrEP_CT_Type_categories= [
            "Qtr PrEP_CT. (Type) Injectables",
            "Qtr PrEP_CT. (Type) Oral",
            "Qtr PrEP_CT. (Type) Other PrEP Type"
        ]

        PrEP_CT_Receiving_PrEP_categories = [
            "Qtr PrEP_CT. Receiving PrEP 15-19yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 15-19yrs, Male",
            "Qtr PrEP_CT. Receiving PrEP 20-24yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 20-24yrs, Male",
            "Qtr PrEP_CT. Receiving PrEP 25-29yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 25-29yrs, Male",
            "Qtr PrEP_CT. Receiving PrEP 30-34yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 30-34yrs, Male",
            "Qtr PrEP_CT. Receiving PrEP 35-39yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 35-39yrs, Male",
            "Qtr PrEP_CT. Receiving PrEP 40-44yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 40-44yrs, Male",
            "Qtr PrEP_CT. Receiving PrEP 45-49 yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 45-49 yrs, Male",
            "Qtr PrEP_CT. Receiving PrEP 50+yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 50+yrs, Male",
        ]


        PrEP_CT_Receiving_PrEP_female_categories = [
            "Qtr PrEP_CT. Receiving PrEP 15-19yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 20-24yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 25-29yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 30-34yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 35-39yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 40-44yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 45-49 yrs, Female",
            "Qtr PrEP_CT. Receiving PrEP 50+yrs, Female",
        ]


        TB_STAT_D_categories =[
          "Qtr TB_STAT_D <1yrs, Female",
          "Qtr TB_STAT_D <1yrs, Male",
          "Qtr TB_STAT_D 1-4yrs, Female",
          "Qtr TB_STAT_D 1-4yrs, Male", 
          "Qtr TB_STAT_D 5-9yrs, Female",
          "Qtr TB_STAT_D 5-9yrs, Male",
          "Qtr TB_STAT_D 10-14yrs, Female",
          "Qtr TB_STAT_D 10-14yrs, Male",
          "Qtr TB_STAT_D 15-19yrs, Female",
          "Qtr TB_STAT_D 15-19yrs, Male",
          "Qtr TB_STAT_D 20-24yrs, Female",
          "Qtr TB_STAT_D 20-24yrs, Male",
          "Qtr TB_STAT_D 25-29yrs, Female",
          "Qtr TB_STAT_D 25-29yrs, Male",
          "Qtr TB_STAT_D 30-34yrs, Female",
          "Qtr TB_STAT_D 30-34yrs, Male",
          "Qtr TB_STAT_D 35-39yrs, Female",
          "Qtr TB_STAT_D 35-39yrs, Male",
           "Qtr TB_STAT_D 40-44yrs, Female",
           "Qtr TB_STAT_D 40-44yrs, Male",
           "Qtr TB_STAT_D 45-49 yrs, Female",
           "Qtr TB_STAT_D 45-49 yrs, Male",
           "Qtr TB_STAT_D 50+yrs, Female",
           "Qtr TB_STAT_D 50+yrs, Male",
        ]

        TB_STAT_N_categories = [
            "Qtr TB_STAT_N Known Positive, Female, 1-4yrs",
            "Qtr TB_STAT_N Known Positive, Female, 10-14yrs",
            "Qtr TB_STAT_N Known Positive, Female, 15-19yrs",
            "Qtr TB_STAT_N Known Positive, Female, 20-24yrs",
            "Qtr TB_STAT_N Known Positive, Female, 25-29yrs",
            "Qtr TB_STAT_N Known Positive, Female, 30-34yrs",
            "Qtr TB_STAT_N Known Positive, Female, 35-39yrs",
            "Qtr TB_STAT_N Known Positive, Female, 35-49yrs",
            "Qtr TB_STAT_N Known Positive, Female, 40-44yrs",
            "Qtr TB_STAT_N Known Positive, Female, 45-49 yrs",
            "Qtr TB_STAT_N Known Positive, Female, 5-9yrs",
            "Qtr TB_STAT_N Known Positive, Female, 50+yrs",
            "Qtr TB_STAT_N Known Positive, Female, <1yrs",
            "Qtr TB_STAT_N Known Positive, Male, 1-4yrs",
            "Qtr TB_STAT_N Known Positive, Male, 10-14yrs",
            "Qtr TB_STAT_N Known Positive, Male, 15-19yrs",
            "Qtr TB_STAT_N Known Positive, Male, 20-24yrs",
            "Qtr TB_STAT_N Known Positive, Male, 25-29yrs",
            "Qtr TB_STAT_N Known Positive, Male, 30-34yrs",
            "Qtr TB_STAT_N Known Positive, Male, 35-39yrs",
            "Qtr TB_STAT_N Known Positive, Male, 40-44yrs",
            "Qtr TB_STAT_N Known Positive, Male, 45-49 yrs",
            "Qtr TB_STAT_N Known Positive, Male, 5-9yrs",
            "Qtr TB_STAT_N Known Positive, Male, 50+yrs",
            "Qtr TB_STAT_N Known Positive, Male, <1yrs",
            "Qtr TB_STAT_N New negative, Female, 1-4yrs",
            "Qtr TB_STAT_N New negative, Female, 10-14yrs",
            "Qtr TB_STAT_N New negative, Female, 15-19yrs",
            "Qtr TB_STAT_N New negative, Female, 20-24yrs",
            "Qtr TB_STAT_N New negative, Female, 25-29yrs",
            "Qtr TB_STAT_N New negative, Female, 30-34yrs",
            "Qtr TB_STAT_N New negative, Female, 35-39yrs",
            "Qtr TB_STAT_N New negative, Female, 40-44yrs",
            "Qtr TB_STAT_N New negative, Female, 45-49 yrs",
            "Qtr TB_STAT_N New negative, Female, 5-9yrs",
            "Qtr TB_STAT_N New negative, Female, 50+yrs",
            "Qtr TB_STAT_N New negative, Female, <1yrs",
            "Qtr TB_STAT_N New negative, Male, 1-4yrs",
            "Qtr TB_STAT_N New negative, Male, 10-14yrs",
            "Qtr TB_STAT_N New negative, Male, 15-19yrs",
            "Qtr TB_STAT_N New negative, Male, 20-24yrs",
            "Qtr TB_STAT_N New negative, Male, 25-29yrs",
            "Qtr TB_STAT_N New negative, Male, 30-34yrs",
            "Qtr TB_STAT_N New negative, Male, 35-39yrs",
            "Qtr TB_STAT_N New negative, Male, 40-44yrs",
            "Qtr TB_STAT_N New negative, Male, 45-49 yrs",
            "Qtr TB_STAT_N New negative, Male, 5-9yrs",
            "Qtr TB_STAT_N New negative, Male, 50+yrs",
            "Qtr TB_STAT_N New negative, Male, <1yrs",
            "Qtr TB_STAT_N New positive, Female, 1-4yrs",
            "Qtr TB_STAT_N New positive, Female, 10-14yrs",
            "Qtr TB_STAT_N New positive, Female, 15-19yrs",
            "Qtr TB_STAT_N New positive, Female, 20-24yrs",
            "Qtr TB_STAT_N New positive, Female, 25-29yrs",
            "Qtr TB_STAT_N New positive, Female, 30-34yrs",
            "Qtr TB_STAT_N New positive, Female, 35-39yrs",
            "Qtr TB_STAT_N New positive, Female, 40-44yrs",
            "Qtr TB_STAT_N New positive, Female, 45-49 yrs",
            "Qtr TB_STAT_N New positive, Female, 5-9yrs",
            "Qtr TB_STAT_N New positive, Female, 50+yrs",
            "Qtr TB_STAT_N New positive, Female, <1yrs",
            "Qtr TB_STAT_N New positive, Male, 1-4yrs",
            "Qtr TB_STAT_N New positive, Male, 10-14yrs",
            "Qtr TB_STAT_N New positive, Male, 15-19yrs",
            "Qtr TB_STAT_N New positive, Male, 20-24yrs",
            "Qtr TB_STAT_N New positive, Male, 25-29yrs",
            "Qtr TB_STAT_N New positive, Male, 30-34yrs",
            "Qtr TB_STAT_N New positive, Male, 35-39yrs",
            "Qtr TB_STAT_N New positive, Male, 40-44yrs",
            "Qtr TB_STAT_N New positive, Male, 45-49 yrs",
            "Qtr TB_STAT_N New positive, Male, 5-9yrs",
            "Qtr TB_STAT_N New positive, Male, 50+yrs",
            "Qtr TB_STAT_N New positive, Male, <1yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 1-4yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 10-14yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 15-19yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 20-24yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 25-29yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 30-34yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 35-39yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 40-44yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 45-49 yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 5-9yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, 50+yrs",
            "Qtr TB_STAT_N Recently tested negative, Female, <1yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 1-4yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 10-14yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 15-19yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 20-24yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 25-29yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 30-34yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 35-39yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 40-44yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 45-49 yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 5-9yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, 50+yrs",
            "Qtr TB_STAT_N Recently tested negative, Male, <1yrs",
        ]

        TX_RTT_categories =[
            "Qtr TX_RTT CD4 <200, 10-14yrs, Female",
            "Qtr TX_RTT CD4 <200, 10-14yrs, Male",
            "Qtr TX_RTT CD4 <200, 15-19yrs, Female",
            "Qtr TX_RTT CD4 <200, 15-19yrs, Male",
            "Qtr TX_RTT CD4 <200, 20-24yrs, Female",
            "Qtr TX_RTT CD4 <200, 20-24yrs, Male",
            "Qtr TX_RTT CD4 <200, 25-29yrs, Female",
            "Qtr TX_RTT CD4 <200, 25-29yrs, Male",
            "Qtr TX_RTT CD4 <200, 30-34yrs, Female",
            "Qtr TX_RTT CD4 <200, 30-34yrs, Male",
            "Qtr TX_RTT CD4 <200, 35-39yrs, Female",
            "Qtr TX_RTT CD4 <200, 35-39yrs, Male",
            "Qtr TX_RTT CD4 <200, 40-44yrs, Female",
            "Qtr TX_RTT CD4 <200, 40-44yrs, Male",
            "Qtr TX_RTT CD4 <200, 45-49 yrs, Female",
            "Qtr TX_RTT CD4 <200, 45-49 yrs, Male",
            "Qtr TX_RTT CD4 <200, 5-9yrs, Female",
            "Qtr TX_RTT CD4 <200, 5-9yrs, Male",
            "Qtr TX_RTT CD4 <200, 50-54yrs, Female",
            "Qtr TX_RTT CD4 <200, 50-54yrs, Male",
            "Qtr TX_RTT CD4 <200, 55-59yrs, Female",
            "Qtr TX_RTT CD4 <200, 55-59yrs, Male",
            "Qtr TX_RTT CD4 <200, 60-64yrs, Female",
            "Qtr TX_RTT CD4 <200, 60-64yrs, Male",
            "Qtr TX_RTT CD4 <200, 65+ yrs, Female",
            "Qtr TX_RTT CD4 <200, 65+ yrs, Male",
            "Qtr TX_RTT CD4 >200, 10-14yrs, Female",
            "Qtr TX_RTT CD4 >200, 10-14yrs, Male",
            "Qtr TX_RTT CD4 >200, 15-19yrs, Female",
            "Qtr TX_RTT CD4 >200, 15-19yrs, Male",
            "Qtr TX_RTT CD4 >200, 20-24yrs, Female",
            "Qtr TX_RTT CD4 >200, 20-24yrs, Male",
            "Qtr TX_RTT CD4 >200, 25-29yrs, Female",
            "Qtr TX_RTT CD4 >200, 25-29yrs, Male",
            "Qtr TX_RTT CD4 >200, 30-34yrs, Female",
            "Qtr TX_RTT CD4 >200, 30-34yrs, Male",
            "Qtr TX_RTT CD4 >200, 35-39yrs, Female",
            "Qtr TX_RTT CD4 >200, 35-39yrs, Male",
            "Qtr TX_RTT CD4 >200, 40-44yrs, Female",
            "Qtr TX_RTT CD4 >200, 40-44yrs, Male",
            "Qtr TX_RTT CD4 >200, 45-49 yrs, Female",
            "Qtr TX_RTT CD4 >200, 45-49 yrs, Male",
            "Qtr TX_RTT CD4 >200, 5-9yrs, Female",
            "Qtr TX_RTT CD4 >200, 5-9yrs, Male",
            "Qtr TX_RTT CD4 >200, 50-54yrs, Female",
            "Qtr TX_RTT CD4 >200, 50-54yrs, Male",
            "Qtr TX_RTT CD4 >200, 55-59yrs, Female"
            ,"Qtr TX_RTT CD4 >200, 55-59yrs, Male",
            "Qtr TX_RTT CD4 >200, 60-64yrs, Female",
            "Qtr TX_RTT CD4 >200, 60-64yrs, Male",
            "Qtr TX_RTT CD4 >200, 65+ yrs, Female",
            "Qtr TX_RTT CD4 >200, 65+ yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 1-4yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 1-4yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 10-14yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 10-14yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 15-19yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 15-19yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 20-24yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 20-24yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 25-29yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 25-29yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 30-34yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 30-34yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 35-39yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 35-39yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 40-44yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 40-44yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 45-49 yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 45-49 yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 5-9yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 5-9yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 50-54yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 50-54yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 55-59yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 55-59yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 60-64yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 60-64yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, 65+ yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, 65+ yrs, Male",
            "Qtr TX_RTT Not eligible for CD4, <1yrs, Female",
            "Qtr TX_RTT Not eligible for CD4, <1yrs, Male",
            "Qtr TX_RTT Unknown CD4, 1-4yrs, Female","Qtr TX_RTT Unknown CD4, 1-4yrs, Male","Qtr TX_RTT Unknown CD4, 10-14yrs, Female","Qtr TX_RTT Unknown CD4, 10-14yrs, Male","Qtr TX_RTT Unknown CD4, 15-19yrs, Female","Qtr TX_RTT Unknown CD4, 15-19yrs, Male","Qtr TX_RTT Unknown CD4, 20-24yrs, Female","Qtr TX_RTT Unknown CD4, 20-24yrs, Male","Qtr TX_RTT Unknown CD4, 25-29yrs, Female","Qtr TX_RTT Unknown CD4, 25-29yrs, Male","Qtr TX_RTT Unknown CD4, 30-34yrs, Female","Qtr TX_RTT Unknown CD4, 30-34yrs, Male","Qtr TX_RTT Unknown CD4, 35-39yrs, Female","Qtr TX_RTT Unknown CD4, 35-39yrs, Male","Qtr TX_RTT Unknown CD4, 40-44yrs, Female","Qtr TX_RTT Unknown CD4, 40-44yrs, Male","Qtr TX_RTT Unknown CD4, 45-49 yrs, Female","Qtr TX_RTT Unknown CD4, 45-49 yrs, Male","Qtr TX_RTT Unknown CD4, 5-9yrs, Female","Qtr TX_RTT Unknown CD4, 5-9yrs, Male","Qtr TX_RTT Unknown CD4, 50-54yrs, Female","Qtr TX_RTT Unknown CD4, 50-54yrs, Male","Qtr TX_RTT Unknown CD4, 55-59yrs, Female","Qtr TX_RTT Unknown CD4, 55-59yrs, Male","Qtr TX_RTT Unknown CD4, 60-64yrs, Female","Qtr TX_RTT Unknown CD4, 60-64yrs, Male","Qtr TX_RTT Unknown CD4, 65+ yrs, Female","Qtr TX_RTT Unknown CD4, 65+ yrs, Male","Qtr TX_RTT Unknown CD4, <1yrs, Female",
            "Qtr TX_RTT Unknown CD4, <1yrs, Male",
        ]


        TX_RTT_IIT_Duration_categories =[
            "Qtr TX_RTT (IIT Duration) IIT <3months",
            "Qtr TX_RTT (IIT Duration) IIT 3-5months",
            "Qtr TX_RTT (IIT Duration) IIT 6+months",

        ]

        Semiannual_CXCA_SCRN_Positve_and_suspected_categories =[
            "Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 10-14yrs",
            "Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 15-19yrs",
            "Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 20-24yrs",
            "Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 25-29yrs",
            "Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 30-34yrs","Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 35-39yrs","Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 40-44yrs","Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 45-49 yrs","Semi-Annual CXCA_SCRN 1st time screened, Positive, Female, 50+yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 10-14yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 15-19yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 20-24yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 25-29yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 30-34yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 35-39yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 40-44yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 45-49 yrs","Semi-Annual CXCA_SCRN 1st time screened, Suspected Cancer, Female, 50+yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 10-14yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 15-19yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 20-24yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 25-29yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 30-34yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 35-39yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 40-44yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 45-49 yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Negative, Female, 50+yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 10-14yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 15-19yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 20-24yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 25-29yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 30-34yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 35-39yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 40-44yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 45-49 yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Positive, Female, 50+yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 10-14yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 15-19yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 20-24yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 25-29yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 30-34yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 35-39yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 40-44yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 45-49 yrs","Semi-Annual CXCA_SCRN Post-treatment follow-up, Suspected Cancer, Female, 50+yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 10-14yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 15-19yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 20-24yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 25-29yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 30-34yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 35-39yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 40-44yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 45-49 yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Negative, Female, 50+yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 10-14yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 15-19yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 20-24yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 25-29yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 30-34yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 35-39yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 40-44yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 45-49 yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Positive, Female, 50+yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 10-14yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 15-19yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 20-24yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 25-29yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 30-34yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 35-39yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 40-44yrs","Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 45-49 yrs",
            "Semi-Annual CXCA_SCRN Rescreened after previous negative, Suspected Cancer, Female, 50+yrs"
        ]

        Semiannual_CXCA_TX_categories =[
            "Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 10-14yrs","Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 15-19yrs",
            "Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 20-24yrs","Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 25-29yrs",
            "Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 30-34yrs","Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 35-39yrs",
            "Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 40-44yrs","Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 45-49 yrs",
            "Semi-Annual CXCA_TX 1st time screened, Cryotherapy, Female, 50+yrs","Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 10-14yrs",
            "Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 15-19yrs","Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 20-24yrs",
            "Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 25-29yrs","Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 30-34yrs",
            "Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 35-39yrs","Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 40-44yrs",
            "Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 45-49 yrs","Semi-Annual CXCA_TX 1st time screened, LEEP, Female, 50+yrs",
            "Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 10-14yrs","Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 15-19yrs",
            "Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 20-24yrs","Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 25-29yrs",
            "Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 30-34yrs","Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 35-39yrs",
            "Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 40-44yrs","Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 45-49 yrs",
            "Semi-Annual CXCA_TX 1st time screened, Thermocoagulation, Female, 50+yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 10-14yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 15-19yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 20-24yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 25-29yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 30-34yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 35-39yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 40-44yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 45-49 yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Cryotherapy, Female, 50+yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 10-14yrs","Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 15-19yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 20-24yrs","Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 25-29yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 30-34yrs","Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 35-39yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 40-44yrs","Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 45-49 yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, LEEP, Female, 50+yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 10-14yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 15-19yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 20-24yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 25-29yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 30-34yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 35-39yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 40-44yrs",
            "Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 45-49 yrs","Semi-Annual CXCA_TX Post-treatment follow-up, Thermocoagulation, Female, 50+yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 10-14yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 15-19yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 20-24yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 25-29yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 30-34yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 35-39yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 40-44yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 45-49 yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Cryotherapy, Female, 50+yrs","Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 10-14yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 15-19yrs","Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 20-24yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 25-29yrs","Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 30-34yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 35-39yrs","Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 40-44yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 45-49 yrs","Semi-Annual CXCA_TX Rescreened after previous negative, LEEP, Female, 50+yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 10-14yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 15-19yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 20-24yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 25-29yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 30-34yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 35-39yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 40-44yrs","Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 45-49 yrs",
            "Semi-Annual CXCA_TX Rescreened after previous negative, Thermocoagulation, Female, 50+yrs",

        ]

        POST_RESP_Sexual_violence_categories = [
            "Semi-Annual POST-RESP. Sexual Violence 10-14yrs, Female","Semi-Annual POST-RESP. Sexual Violence 10-14yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 15-19yrs, Female","Semi-Annual POST-RESP. Sexual Violence 15-19yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 20-24yrs, Female","Semi-Annual POST-RESP. Sexual Violence 20-24yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 25-29yrs, Female","Semi-Annual POST-RESP. Sexual Violence 25-29yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 30-34yrs, Female","Semi-Annual POST-RESP. Sexual Violence 30-34yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 35-39yrs, Female","Semi-Annual POST-RESP. Sexual Violence 35-39yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 40-44yrs, Female","Semi-Annual POST-RESP. Sexual Violence 40-44yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 45-49 yrs, Female","Semi-Annual POST-RESP. Sexual Violence 45-49 yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence 50+yrs, Female","Semi-Annual POST-RESP. Sexual Violence 50+yrs, Male",
            "Semi-Annual POST-RESP. Sexual Violence <10yrs, Female","Semi-Annual POST-RESP. Sexual Violence <10yrs, Male"
        ]


        POST_RESP_Sexual_violence_with_PEP_categories = [
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 10-14yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 10-14yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 15-19yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 15-19yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 20-24yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 20-24yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 25-29yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 25-29yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 30-34yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 30-34yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 35-39yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 35-39yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 40-44yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 40-44yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 45-49 yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 45-49 yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 50+yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) 50+yrs, Male",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) <10yrs, Female",
            "Semi-Annual POST-RESP. Number of People Receiving Post-Exposure Prophylaxis (PEP) <10yrs, Male"
        ]

        TB_PREV_D_categories =["Semi-Annual TB_PREV_D Already on ART, Female, 0-14 yrs",
                               "Semi-Annual TB_PREV_D Already on ART, Female, 15+yrs",
                               "Semi-Annual TB_PREV_D Already on ART, Male, 0-14 yrs",
                               "Semi-Annual TB_PREV_D Already on ART, Male, 15+yrs",
                               "Semi-Annual TB_PREV_D New on ART, Female, 0-14 yrs",
                               "Semi-Annual TB_PREV_D New on ART, Female, 15+yrs",
                               "Semi-Annual TB_PREV_D New on ART, Male, 0-14 yrs",
                               "Semi-Annual TB_PREV_D New on ART, Male, 15+yrs",
        ]


        TB_PREV_N_categories=[
            "Semi-Annual TB_PREV_N Already on ART, Female, 0-14 yrs",
            "Semi-Annual TB_PREV_N Already on ART, Female, 15+yrs",
            "Semi-Annual TB_PREV_N Already on ART, Male, 0-14 yrs",
            "Semi-Annual TB_PREV_N Already on ART, Male, 15+yrs",
            "Semi-Annual TB_PREV_N New on ART, Female, 0-14 yrs",
            "Semi-Annual TB_PREV_N New on ART, Female, 15+yrs",
            "Semi-Annual TB_PREV_N New on ART, Male, 0-14 yrs",
            "Semi-Annual TB_PREV_N New on ART, Male, 15+yrs"
        ]

        # Calculate sums for each group, handling potential errors, and filtering by Period and OrgUnit
        def calculate_sum(df, categories, period, orgunit):
            total_sum = 0
            for category in categories:
                try:
                    df_filtered = df[(df['Indicator'] == category) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)]
                    if not df_filtered.empty:
                        value_series = df_filtered['Value']
                        if pd.api.types.is_numeric_dtype(value_series):
                            total_sum += value_series.sum()
                        else:
                            print(f"Warning: Non-numeric value found in 'Value' column for category '{category}', Period '{period}', OrgUnit '{orgunit}'. Skipping summation for this category.")
                            return None
                    else:
                        pass # No matching rows for this category, period and orgunit
                except KeyError as e:
                    print(f"Error: Column not found: {e}. Check your Excel file's column names.")
                    return None # Return None to indicate an error
                except TypeError as e:
                    print(f"Error: {e}. Check the data type of your 'Value' column. It should be numeric.")
                    return None
            return total_sum

        # Create a new column "Validation Result"
        df['Validation Result'] = ''  # Initialize the column

        # --- Validation for HTS SELF ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate HTS SELF sums
            hts_self_user_sum = calculate_sum(df, hts_self_user_categories, period, orgunit)
            hts_self_sum = calculate_sum(df, hts_self_categories, period, orgunit)

            if hts_self_user_sum is not None and hts_self_sum is not None and hts_self_user_sum != hts_self_sum:
                # Identify rows belonging to HTS SELF categories for this Period and OrgUnit
                hts_rows_mask = (df['Indicator'].isin(hts_self_categories + hts_self_user_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[hts_rows_mask, 'Validation Result'] = 'HTS SELF Unassisted is not equal to TestUser'
            elif hts_self_user_sum is None or hts_self_sum is None:
                hts_rows_mask = (df['Indicator'].isin(hts_self_categories + hts_self_user_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[hts_rows_mask, 'Validation Result'] = '#N/A'

        # --- Validation for PrEP CT & Distribution ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate PrEP CT sums
            PrEP_CT_Distribution = calculate_sum(df, PrEP_CT_Distribution_categories, period, orgunit)
            PrEP_CT_Receiving_PrEP = calculate_sum(df, PrEP_CT_Receiving_PrEP_categories, period, orgunit)

            if PrEP_CT_Distribution is not None and PrEP_CT_Receiving_PrEP is not None and PrEP_CT_Distribution != PrEP_CT_Receiving_PrEP:
                # Identify rows belonging to PrEP CT categories for this Period and OrgUnit
                prep_rows_mask = (df['Indicator'].isin(PrEP_CT_Distribution_categories + PrEP_CT_Receiving_PrEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[prep_rows_mask, 'Validation Result'] = 'PrEP_CT. Receiving PrEP is not equal to PrEP_CT. (Distribution)'
            elif PrEP_CT_Distribution is None or PrEP_CT_Receiving_PrEP is None:
                prep_rows_mask = (df['Indicator'].isin(PrEP_CT_Distribution_categories + PrEP_CT_Receiving_PrEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[prep_rows_mask, 'Validation Result'] = '#N/A'

        # --- Validation for PrEP CT & Test Result ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate PrEP CT sums
            PrEP_CT_TestResult = calculate_sum(df, PrEP_CT_TestResult_categories, period, orgunit)
            PrEP_CT_Receiving_PrEP = calculate_sum(df, PrEP_CT_Receiving_PrEP_categories, period, orgunit)

            if PrEP_CT_TestResult is not None and PrEP_CT_Receiving_PrEP is not None and PrEP_CT_TestResult != PrEP_CT_Receiving_PrEP:
                # Identify rows belonging to PrEP CT categories for this Period and OrgUnit
                prepp_rows_mask = (df['Indicator'].isin(PrEP_CT_TestResult_categories + PrEP_CT_Receiving_PrEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[prepp_rows_mask, 'Validation Result'] = 'PrEP_CT. Receiving PrEP is not equal to PrEP_CT. (Testresult)'
            elif PrEP_CT_TestResult is None or PrEP_CT_Receiving_PrEP is None:
                prepp_rows_mask = (df['Indicator'].isin(PrEP_CT_TestResult_categories + PrEP_CT_Receiving_PrEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[prepp_rows_mask, 'Validation Result'] = '#N/A'

        # --- Validation for PrEP CT & Type ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate PrEP CT sums
            PrEP_CT_Type = calculate_sum(df, PrEP_CT_Type_categories, period, orgunit)
            PrEP_CT_Receiving_PrEP = calculate_sum(df, PrEP_CT_Receiving_PrEP_categories, period, orgunit)

            if PrEP_CT_Type is not None and PrEP_CT_Receiving_PrEP is not None and PrEP_CT_Type != PrEP_CT_Receiving_PrEP:
                # Identify rows belonging to PrEP CT categories for this Period and OrgUnit
                preppp_rows_mask = (df['Indicator'].isin(PrEP_CT_Type_categories + PrEP_CT_Receiving_PrEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[preppp_rows_mask, 'Validation Result'] = 'PrEP_CT. Receiving PrEP is not equal to PrEP_CT. (Type)'
            elif PrEP_CT_Type is None or PrEP_CT_Receiving_PrEP is None:
                preppp_rows_mask = (df['Indicator'].isin(PrEP_CT_Type_categories + PrEP_CT_Receiving_PrEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[preppp_rows_mask, 'Validation Result'] = '#N/A'

        # --- Validation for PrEP CT(female only) & PrEP_CT (PregBF) ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate PrEP CT sums
            PrEP_CT_PregBF = calculate_sum(df, PrEP_CT_PregBF_categories, period, orgunit)
            PrEP_CT_Receiving_PrEP_female = calculate_sum(df, PrEP_CT_Receiving_PrEP_female_categories, period, orgunit)

            if PrEP_CT_PregBF is not None and PrEP_CT_Receiving_PrEP_female is not None and PrEP_CT_PregBF > PrEP_CT_Receiving_PrEP_female:
                # Identify rows belonging to PrEP CT categories for this Period and OrgUnit
                prepppp_rows_mask = (df['Indicator'].isin(PrEP_CT_PregBF_categories + PrEP_CT_Receiving_PrEP_female_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[prepppp_rows_mask, 'Validation Result'] = 'PrEP_CT. (PregBF) is greater than PrEP_CT. Receiving PrEP (female only)'
            elif PrEP_CT_PregBF is None or PrEP_CT_Receiving_PrEP_female is None:
                prepppp_rows_mask = (df['Indicator'].isin(PrEP_CT_PregBF_categories + PrEP_CT_Receiving_PrEP_female_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[preppp_rows_mask, 'Validation Result'] = '#N/A'

        # --- Validation for TB_STAT_D & TB_STAT_N ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate TB_STAT sums
            TB_STAT_N = calculate_sum(df, TB_STAT_N_categories, period, orgunit)
            TB_STAT_D = calculate_sum(df, TB_STAT_D_categories, period, orgunit)

            if TB_STAT_N is not None and TB_STAT_D is not None and TB_STAT_N != TB_STAT_D:
                # Identify rows belonging to TB_STAT categories for this Period and OrgUnit
                tb_stat_rows_mask = (df['Indicator'].isin(TB_STAT_N_categories + TB_STAT_D_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[tb_stat_rows_mask, 'Validation Result'] = 'TB_STAT_D is not equal to TB_STAT_N'
            elif TB_STAT_N is None or TB_STAT_D is None:
                tb_stat_rows_mask = (df['Indicator'].isin(TB_STAT_N_categories + TB_STAT_D_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[tb_stat_rows_mask, 'Validation Result'] = '#N/A'


        # --- Validation for TX_RTT & TX_RTT_IIT_Duration ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate TX_RTT sums
            TX_RTT_IIT_Duration = calculate_sum(df, TX_RTT_IIT_Duration_categories, period, orgunit)
            TX_RTT = calculate_sum(df, TX_RTT_categories, period, orgunit)

            if TX_RTT is not None and TX_RTT_IIT_Duration is not None and TX_RTT_IIT_Duration != TX_RTT:
                # Identify rows belonging to TX_RTT categories for this Period and OrgUnit
                tx_rtt_rows_mask = (df['Indicator'].isin(TX_RTT_categories + TX_RTT_IIT_Duration_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[tx_rtt_rows_mask, 'Validation Result'] = 'TX_RTT_IIT_Duration is not equal to TX_RTT'
            elif TX_RTT is None or TX_RTT_IIT_Duration is None:
                tx_rtt_rows_mask = (df['Indicator'].isin(TX_RTT_categories + TX_RTT_IIT_Duration_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[tx_rtt_rows_mask, 'Validation Result'] = '#N/A'


        # --- Validation for Semiannual_CXCA_TX & Semiannual_CXCA_SCRN_Positve_and_suspected ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate CXCA sums
            Semiannual_CXCA_TX = calculate_sum(df, Semiannual_CXCA_TX_categories, period, orgunit)
            Semiannual_CXCA_SCRN_Positve_and_suspected = calculate_sum(df, Semiannual_CXCA_SCRN_Positve_and_suspected_categories, period, orgunit)

            if Semiannual_CXCA_TX is not None and Semiannual_CXCA_SCRN_Positve_and_suspected is not None and Semiannual_CXCA_TX > Semiannual_CXCA_SCRN_Positve_and_suspected:
                # Identify rows belonging to TX_RTT categories for this Period and OrgUnit
                cxca_rows_mask = (df['Indicator'].isin(Semiannual_CXCA_TX_categories + Semiannual_CXCA_SCRN_Positve_and_suspected_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[cxca_rows_mask, 'Validation Result'] = 'CXCA_TX is greater than CXCA_SCRN_Positive_and_suspected'
            elif Semiannual_CXCA_TX is None or Semiannual_CXCA_SCRN_Positve_and_suspected is None:
                cxca_rows_mask = (df['Indicator'].isin(Semiannual_CXCA_TX_categories + Semiannual_CXCA_SCRN_Positve_and_suspected_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[cxca_rows_mask, 'Validation Result'] = '#N/A'


        # --- Validation for POST_RESP_Sexual_violence & POST_RESP_Sexual_violence_with_PEP ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate POST_RESP sums
            POST_RESP_Sexual_violence = calculate_sum(df, POST_RESP_Sexual_violence_categories, period, orgunit)
            POST_RESP_Sexual_violence_with_PEP = calculate_sum(df, POST_RESP_Sexual_violence_with_PEP_categories, period, orgunit)

            if POST_RESP_Sexual_violence is not None and POST_RESP_Sexual_violence_with_PEP is not None and POST_RESP_Sexual_violence_with_PEP > POST_RESP_Sexual_violence:
                # Identify rows belonging to POST_RESP categories for this Period and OrgUnit
                post_resp_rows_mask = (df['Indicator'].isin(POST_RESP_Sexual_violence_categories + POST_RESP_Sexual_violence_with_PEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[post_resp_rows_mask, 'Validation Result'] = 'POST_RESP_Sexual_violence_with_PEP is greater than POST_RESP_Sexual_violence'
            elif POST_RESP_Sexual_violence is None or POST_RESP_Sexual_violence_with_PEP is None:
                post_resp_rows_mask = (df['Indicator'].isin(POST_RESP_Sexual_violence_categories + POST_RESP_Sexual_violence_with_PEP_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[post_resp_rows_mask, 'Validation Result'] = '#N/A'


        # --- Validation for TB_PREV_D & TB_PREV_N ---
        for (period, orgunit), group_df in df.groupby(['Period', 'OrgUnit']):
            # Calculate TB_PREV sums
            TB_PREV_N = calculate_sum(df, TB_PREV_N_categories, period, orgunit)
            TB_PREV_D = calculate_sum(df, TB_PREV_D_categories, period, orgunit)

            if TB_PREV_N is not None and TB_PREV_D is not None and TB_PREV_N > TB_PREV_D:
                # Identify rows belonging to TB_PREV categories for this Period and OrgUnit
                tb_prev_rows_mask = (df['Indicator'].isin(TB_PREV_D_categories + TB_PREV_N_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[tb_prev_rows_mask, 'Validation Result'] = 'TB_PREV_N is greater than TB_PREV_D'
            elif TB_PREV_N is None or TB_PREV_D is None:
                tb_prev_rows_mask = (df['Indicator'].isin(TB_PREV_D_categories + TB_PREV_N_categories)) & (df['Period'] == period) & (df['OrgUnit'] == orgunit)
                df.loc[tb_prev_rows_mask, 'Validation Result'] = '#N/A'

        # Save the modified DataFrame to a new Excel file
        df.to_excel(output_file_path, index=False)
        print(f"Validation complete. Results saved to {output_file_path}")

    except FileNotFoundError:
        print(f"Error: File not found at {excel_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
        print("Please ensure the Excel file is not open and that the column names are correct.")

if __name__ == "__main__":
    # Specify the paths to your input and output Excel files
    input_excel_file = "C:/Users/DELL/Documents/DataFi/DATIM Optional Data Element Codelist- Structure (1).xlsx"
    output_excel_file = "C:/Users/DELL/Documents/DataFi/DATIM Optional Data Element Codelist- Structure with validation rules.xlsx"  

    # Run the validation
    validate_hts_prep_data_separate(input_excel_file, output_excel_file)
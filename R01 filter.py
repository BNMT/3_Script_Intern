import pandas as pd
import numpy as np


# apply file
df = pd.ExcelFile(r"C:/Users/VivoBook/Downloads/R01_Báo cáo danh mục nhân viên_Assignment Employee Report Full_V1.xls")
# sheet 1
df_goc = pd.read_excel(df, sheet_name='gốc', skiprows=15)

# sheet 2
df_check_0 = pd.read_excel(df, sheet_name='check', skiprows=15)
df_check = df_check_0[['Seq', 'Employee ID', 'Employee Full Name', 'e.full name-check', 'Local Name', 'local name-check', 'industry-check', 'GF Industry-check', 'Legal Entity', 'Business Unit', 'dep-check', 'dep vn-check', 'w.location-check', 'j.code-check', 'job-check', 'job vn-check', 'w.email-check ', 'w.phone-check', 'LM-check', 'LM assi-check', 'LM name-check', 'contract type-check', 'contract start date-check', 'contract end date-check', 'contract num-check', 'dob-check', 'gender-check', 'country-check', 'town-check', 'edu-check', 'marry-check', 'ethinicity-check', 'religion-check', 'residential-check', 'street-check', 'district-check', 'city-check', 'province-check', 'address-check', 'temp street-check', 'temp district-check', 'temp city-check', 'temp address check', 'id check', 'issue date-check', 'place of issue-check', 'nation id-check', 'bank acc-check', 'branch-check', 'bank name-check']]

# sheet 3
df = df_check_0.copy()
df['e.full name-check'] = np.where(df['e.full name-check'] == 'dư khoảng trắng', df['Business Unit'], df['e.full name-check'])
df['local name-check'] = np.where(df['local name-check'] == ('dư khoảng trắng' or 'thiếu'), df['Business Unit'], df['local name-check'])

for col in df.columns:
    df[col] = np.where(df[col] == 'thiếu', df['Business Unit'], df[col])

df_count = df[['e.full name-check', 'local name-check', 'industry-check', 'GF Industry-check', 'dep-check', 'dep vn-check', 'w.location-check', 'j.code-check', 'job-check', 'job vn-check', 'w.email-check ', 'w.phone-check', 'LM-check', 'LM assi-check', 'LM name-check', 'contract type-check', 'contract start date-check', 'contract num-check', 'dob-check', 'gender-check', 'country-check', 'town-check', 'edu-check', 'marry-check', 'ethinicity-check', 'religion-check', 'residential-check', 'street-check', 'district-check', 'city-check', 'province-check', 'address-check', 'id check', 'issue date-check', 'place of issue-check', 'nation id-check', 'bank acc-check', 'branch-check', 'bank name-check']]
pivot = []
for col in df_count.columns:
    x = df_count[col].value_counts()
    print(x)
    print(f'>>> Total value of {col} = ', sum(x), '\n\n')
    pivot.append(x)
df_pivot = pd.DataFrame(pivot,
                        index=['Employee Full Name', 'Local Name', 'Industry', 'GF Industry', 'Department', 'Department VN', 'Work Location', 'Job Code', 'Job', 'Job VN', 'Work Email', 'Working Phone', 'Line Manager ID', 'Line manager Assignment', 'Line Manager Name', 'Labor Contract Type', 'Labor Contract Start Date', 'Contract number', 'Date Of Birth', 'Gender', 'Country Of Birth', 'Town of Birth', 'Highest Education Level', 'Marital Status', 'Ethnicity', 'Religion', 'Residential Type', 'Permanent Stress', 'Permanent District or Town', 'Permanent City', 'Permanent Province', 'Permanent Residential Address', 'National ID', 'Issue Date', 'Place of Issue', 'National ID Type', 'Bank Account', 'Branch', 'Bank Name'],
                        columns=['Aquafeed', 'Binh Dinh', 'Cambodia', 'Cambodia Farm', 'Cambodia Feed', 'Central region Farm', 'D&F', 'Dong Nai', 'Excel-Tech', 'Farm Functional Department - North & Central', 'Farm Functional Departments - South', 'Farm Functional Units', 'Feddy', 'FedFarm', 'Feed Functional Units', 'Feed VN - General Management', 'Feed-Aqua VN', 'GREENIFIQUE', 'Ha Nam', 'Head Office', 'Hung Yen', 'Laos', 'LEBOUCHER', 'Logs - QD Trans', 'Logs - QD Trans - Operations', 'Long An', 'Long An - Aquafeed Functional Office', 'Myanmar', 'Northern 1', 'Northern 2', 'Northern Poultry VN', 'PREMIX', 'Southern 1', 'Southern 2', 'Southern Poultry VN', 'Tek - QDTek MB', 'Tek - QDTek MN', 'Viet Tho', 'Vinh Long']).transpose()

# sheet 4+
df_aquafeed = df_check.loc[df_check['Business Unit'] == 'Aquafeed']
df_binhdinh = df_check.loc[df_check['Business Unit'] == 'Binh Dinh']
df_cambodia = df_check.loc[df_check['Business Unit'] == 'Cambodia']
df_cambodiafarm = df_check.loc[df_check['Business Unit'] == 'Cambodia Farm']
df_cambodiafeed = df_check.loc[df_check['Business Unit'] == 'Cambodia Feed']
df_centralregionfarm = df_check.loc[df_check['Business Unit'] == 'Central region Farm']
df_dnf = df_check.loc[df_check['Business Unit'] == 'D&F']
df_dongnai = df_check.loc[df_check['Business Unit'] == 'Dong Nai']
df_exceltech = df_check.loc[df_check['Business Unit'] == 'Excel-Tech']
df_farmfunctionaldepartmentnorthcentral = df_check.loc[df_check['Business Unit'] == 'Farm Functional Department - North & Central']
df_farmfunctionaldepartmentssouth = df_check.loc[df_check['Business Unit'] == 'Farm Functional Departments - South']
df_farmfunctionalunits = df_check.loc[df_check['Business Unit'] == 'Farm Functional Units']
df_feddy = df_check.loc[df_check['Business Unit'] == 'Feddy']
df_fedfarm = df_check.loc[df_check['Business Unit'] == 'FedFarm']
df_feedfunctionalunits = df_check.loc[df_check['Business Unit'] == 'Feed Functional Units']
df_feedvngeneralmanagement = df_check.loc[df_check['Business Unit'] == 'Feed VN - General Management']
df_feedaquavn = df_check.loc[df_check['Business Unit'] == 'Feed-Aqua VN']
df_greenifique = df_check.loc[df_check['Business Unit'] == 'GREENIFIQUE']
df_hanam = df_check.loc[df_check['Business Unit'] == 'Ha Nam']
df_headoffice = df_check.loc[df_check['Business Unit'] == 'Head Office']
df_hungyen = df_check.loc[df_check['Business Unit'] == 'Hung Yen']
df_laos = df_check.loc[df_check['Business Unit'] == 'Laos']
df_leboucher = df_check.loc[df_check['Business Unit'] == 'LEBOUCHER']
df_logsqdtrans = df_check.loc[df_check['Business Unit'] == 'Logs - QD Trans']
df_logsqdtransoperations = df_check.loc[df_check['Business Unit'] == 'Logs - QD Trans - Operations']
df_longan = df_check.loc[df_check['Business Unit'] == 'Long An']
df_longanaquafeedfunctionaloffice = df_check.loc[df_check['Business Unit'] == 'Long An - Aquafeed Functional Office']
df_myanmar = df_check.loc[df_check['Business Unit'] == 'Myanmar']
df_northern1 = df_check.loc[df_check['Business Unit'] == 'Northern 1']
df_northern2 = df_check.loc[df_check['Business Unit'] == 'Northern 2']
df_northernpoultryvn = df_check.loc[df_check['Business Unit'] == 'Northern Poultry VN']
df_premix = df_check.loc[df_check['Business Unit'] == 'PREMIX']
df_southern1 = df_check.loc[df_check['Business Unit'] == 'Southern 1']
df_southern2 = df_check.loc[df_check['Business Unit'] == 'Southern 2']
df_southernpoultryvn = df_check.loc[df_check['Business Unit'] == 'Southern Poultry VN']
df_tekqdtekmb = df_check.loc[df_check['Business Unit'] == 'Tek - QDTek MB']
df_tekqdtekmn = df_check.loc[df_check['Business Unit'] == 'Tek - QDTek MN']
df_viettho = df_check.loc[df_check['Business Unit'] == 'Viet Tho']
df_vinhlong = df_check.loc[df_check['Business Unit'] == 'Vinh Long']






with pd.ExcelWriter('R01_Báo cáo danh mục nhân viên.xlsx') as writer:
    # sheet 1
    df_goc.to_excel(writer, sheet_name='gốc', index=False)
    # sheet 2
    df_check.to_excel(writer, sheet_name='check', index=False)
    # sheet 3
    df_pivot.to_excel(writer, sheet_name='pivot')
    # sheet 4+
    df_aquafeed.to_excel(writer, sheet_name='Aquafeed', index=False)
    df_binhdinh.to_excel(writer, sheet_name='Binh Dinh', index=False)
    df_cambodia.to_excel(writer, sheet_name='Cambodia', index=False)
    df_cambodiafarm.to_excel(writer, sheet_name='Cambodia Farm', index=False)
    df_cambodiafeed.to_excel(writer, sheet_name='Cambodia Feed', index=False)
    df_centralregionfarm.to_excel(writer, sheet_name='Central region Farm', index=False)
    df_dnf.to_excel(writer, sheet_name='D&F', index=False)
    df_dongnai.to_excel(writer, sheet_name='Dong Nai', index=False)
    df_exceltech.to_excel(writer, sheet_name='Excel-Tech', index=False)
    df_farmfunctionaldepartmentnorthcentral.to_excel(writer, sheet_name='Farm Functional Department - North & Central', index=False)
    df_farmfunctionaldepartmentssouth.to_excel(writer, sheet_name='Farm Functional Departments - South', index=False)
    df_farmfunctionalunits.to_excel(writer, sheet_name='Farm Functional Units', index=False)
    df_feddy.to_excel(writer, sheet_name='Feddy', index=False)
    df_fedfarm.to_excel(writer, sheet_name='FedFarm', index=False)
    df_feedfunctionalunits.to_excel(writer, sheet_name='Feed Functional Units', index=False)
    df_feedvngeneralmanagement.to_excel(writer, sheet_name='Feed VN - General Management', index=False)
    df_feedaquavn.to_excel(writer, sheet_name='Feed-Aqua VN', index=False)
    df_greenifique.to_excel(writer, sheet_name='GREENIFIQUE', index=False)
    df_hanam.to_excel(writer, sheet_name='Ha Nam', index=False)
    df_headoffice.to_excel(writer, sheet_name='Head Office', index=False)
    df_hungyen.to_excel(writer, sheet_name='Hung Yen', index=False)
    df_laos.to_excel(writer, sheet_name='Laos', index=False)
    df_leboucher.to_excel(writer, sheet_name='LEBOUCHER', index=False)
    df_logsqdtrans.to_excel(writer, sheet_name='Logs - QD Trans', index=False)
    df_logsqdtransoperations.to_excel(writer, sheet_name='Logs - QD Trans - Operations', index=False)
    df_longan.to_excel(writer, sheet_name='Long An', index=False)
    df_longanaquafeedfunctionaloffice.to_excel(writer, sheet_name='Long An - Aquafeed Functional Office', index=False)
    df_myanmar.to_excel(writer, sheet_name='Myanmar', index=False)
    df_northern1.to_excel(writer, sheet_name='Northern 1', index=False)
    df_northern2.to_excel(writer, sheet_name='Northern 2', index=False)
    df_northernpoultryvn.to_excel(writer, sheet_name='Northern Poultry VN', index=False)
    df_premix.to_excel(writer, sheet_name='PREMIX', index=False)
    df_southern1.to_excel(writer, sheet_name='Southern 1', index=False)
    df_southern2.to_excel(writer, sheet_name='Southern 2', index=False)
    df_southernpoultryvn.to_excel(writer, sheet_name='Southern Poultry VN', index=False)
    df_tekqdtekmb.to_excel(writer, sheet_name='Tek - QDTek MB', index=False)
    df_tekqdtekmn.to_excel(writer, sheet_name='Tek - QDTek MN', index=False)
    df_viettho.to_excel(writer, sheet_name='Viet Tho', index=False)
    df_vinhlong.to_excel(writer, sheet_name='Vinh Long', index=False)

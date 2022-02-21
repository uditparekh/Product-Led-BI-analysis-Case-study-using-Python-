import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from spinner import Spinner
import seaborn as sns


s=Spinner()
s.start()

df = pd.read_excel('BI Analyst Case Study Sample Data.xlsx')


df.columns = df.columns.str.replace(' ','_')

print(df)




'''
Number of user account
'''

Count_User_ID = len(df.User_ID)

print("No. of users: " + str(Count_User_ID))

'''
filling blank column with zero
'''

df['secondsOnLoadingPage'] = df['secondsOnLoadingPage'].fillna(0)


df['isInlineBaComplete'] = df['InlineBaComplete'].apply(lambda x: 'Complete BA' if x == True else 'Incomplete BA')

df['Is_Product_Qualified_Lead?'] = df['Product_Qualified_Lead?'].apply(lambda x: 'Qualified Lead' if x == True else 'Not Qualified Lead')

'''

Converting seconds to minutes for better understanding

'''

def convert(seconds):
    seconds = seconds % (24 * 3600)
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60
    return "%d:%02d:%02d" % (hour, minutes, seconds)


df['Total_time_spent_Min'] = df['totalSecondsSpentOnAssessment'].apply(convert)
df['Load_time_minutes'] = df['secondsOnLoadingPage'].apply(convert)


'''
Analysis based on design version

Plot to check how many account users are using design version
'''

plt.figure(figsize=(10, 8))
l = sns.countplot(x = "Account_Created_On_Design_SS_Version", data = df)
sns.despine()
plt.title("Number of Accounts per Design Version", fontsize = 18)
plt.xlabel("Design Version", fontsize = 14)
plt.ylabel("Count of Account", fontsize = 14)

for p in l.patches:
    l.annotate(p.get_height(), (p.get_x() + p.get_width() / 2., p.get_height() + 10), ha = 'center', va = 'center')

plt.savefig('PI/Count_of_Account_per_version.jpeg', dpi=100)

table = pd.pivot_table(df,
                       index=['Account_Created_On_Design_SS_Version'],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)



'''
Taking as IsInlineBAComplete as major benchmark

lets check number of users as per version

'''

plt.figure(figsize=(10, 8))
l = sns.countplot(x = "Account_Created_On_Design_SS_Version", hue = "isInlineBaComplete", data = df)
sns.despine()
plt.title("Number of Users complete BA", fontsize = 18)
plt.xlabel("Design Version", fontsize = 14)
plt.ylabel("Count of Users", fontsize = 14)

for p in l.patches:
    l.annotate(p.get_height(), (p.get_x() + p.get_width() / 2., p.get_height() + 10), ha = 'center', va = 'center')

plt.savefig('PI/Count_of_Users_InlineBA.jpeg', dpi=100)

table2 = pd.pivot_table(df,
                       index=['Account_Created_On_Design_SS_Version'],
                       columns=["isInlineBaComplete"],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

table2= table2.rename(columns={"User_ID": ' ' })
table2= table2.rename(columns={"isInlineBaComplete": ' Account_Created_On_Design_SS_Version' })
table2= table2.rename(index={'Account_Created_On_Design_SS_Version': ' ' })

'''
Total time spent by hours by Users as per version
'''

table3 = pd.pivot_table(df,
                       index=['Account_Created_On_Design_SS_Version'],
                       columns=["isInlineBaComplete"],
                       values=["totalSecondsSpentOnAssessment"],
                       aggfunc={"totalSecondsSpentOnAssessment": np.sum},
                       fill_value=pd.Timedelta(hours=0),
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

table3_ = table3.applymap(convert)

table3_= table3_.rename(columns={"totalSecondsSpentOnAssessment": ' ' })
table3_= table3_.rename(columns={"isInlineBaComplete": ' Account_Created_On_Design_SS_Version' })
table3_= table3_.rename(index={'Account_Created_On_Design_SS_Version': ' ' })


'''
Dividing the dataframe in InlineBA complete between true and false for further analysis
'''


InlineBA_True = df[(df["InlineBaComplete"] == True)]
InlineBA_False = df[(df["InlineBaComplete"] == False)]

tableSample = pd.pivot_table(df,
                       index=['Account_Created_On_Design_SS_Version'],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins=False)
                       #margins_name="x Grand Total")
                       #fill_value=0)


table4 = pd.pivot_table(InlineBA_True,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins=False)
                       #margins_name="x Grand Total")
                       #fill_value=0)


True_table = table4.div( tableSample.iloc[:,-1], axis=0 )*100.00
True_table =True_table.round(2)
#True_table=True_table[:-1]
#True_table= True_table.rename(columns={"User_ID": ' % of User_ID ' })
#True_table= True_table.rename(columns={"PisInlineBaComplete": ' ' })
#True_table= True_table.rename(index={'Account_Created_On_Design_SS_Version': ' ' })




table5 = pd.pivot_table(InlineBA_False,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins=False)
                       #margins_name="x Grand Total")
                       #fill_value=0)


False_table = table5.div( tableSample.iloc[:,-1], axis=0 )*100.00
False_table =False_table.round(2)


plt.figure(figsize=(14, 8))
l = sns.countplot(x = "Account_Created_On_Design_SS_Version",  data = InlineBA_False)
sns.despine()
plt.title("Conversion rate by percent_false", fontsize = 18)
plt.xlabel("Account_Created_On_Design_SS_Version", fontsize = 14)
plt.ylabel("User_ID", fontsize = 14)

for p in l.patches:
    l.annotate(p.get_height(), (p.get_x() + p.get_width() / 2., p.get_height() + 5), ha = 'center', va = 'center')

plt.savefig('PI/Conversion_rate_false.jpeg', dpi=100)



plt.figure(figsize=(14, 8))
l = sns.countplot(x = "Account_Created_On_Design_SS_Version",  data = InlineBA_True)
sns.despine()
plt.title("Conversion rate by percent_True", fontsize = 18)
plt.xlabel("Account_Created_On_Design_SS_Version", fontsize = 14)
plt.ylabel("User_ID", fontsize = 14)

for p in l.patches:
    l.annotate(p.get_height(), (p.get_x() + p.get_width() / 2., p.get_height() + 5), ha = 'center', va = 'center')

plt.savefig('PI/Conversion_rate_True.jpeg', dpi=100)

#False_table=False_table[:-1]
#False_table= False_table.rename(columns={"User_ID": ' % of User_ID ' })
#False_table= False_table.rename(columns={"PisInlineBaComplete": ' ' })
#False_table= False_table.rename(index={'Account_Created_On_Design_SS_Version': ' ' })


MyLabels = ["DPA 1.2", "DPA 1.3", "DPO 1.0", "DPO 1.1"]
plt.figure(figsize=(16,8))
# plot chart
ax1 = plt.subplot(121, aspect='equal')
table5.plot(kind='pie', y = 'User_ID', ax=ax1, autopct='%1.0f%%',
 startangle=90, shadow=False, labels=MyLabels, legend = False, fontsize=14)
plt.title('User with BA IncompleteComplete ')
plt.axis('equal')

plt.savefig('PI/Pie_false.jpeg', dpi=100)


plt.figure(figsize=(16,8))
# plot chart
ax1 = plt.subplot(121, aspect='equal')
table4.plot(kind='pie', y = 'User_ID', ax=ax1, autopct='%1.0f%%',
 startangle=90, shadow=False, labels=MyLabels, legend = False, fontsize=14)
plt.title('User with BA Complete ')
plt.axis('equal')
plt.savefig('PI/Pie_True.jpeg', dpi=100)

tableSample1 = pd.pivot_table(InlineBA_False,
                       index=['Account_Created_On_Design_SS_Version'],
                       columns=['Is_Product_Qualified_Lead?'],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

tableSample1= tableSample1.rename(columns={"User_ID": ' ' })
tableSample1= tableSample1.rename(columns={"Is_Product_Qualified_Lead": ' Account_Created_On_Design_SS_Version' })
tableSample1= tableSample1.rename(index={'Account_Created_On_Design_SS_Version': ' ' })


tableSample2 = pd.pivot_table(InlineBA_True,
                       index=['Account_Created_On_Design_SS_Version'],
                       columns=['Is_Product_Qualified_Lead?'],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

tableSample2= tableSample2.rename(columns={"User_ID": ' ' })
tableSample2= tableSample2.rename(columns={"Is_Product_Qualified_Lead": ' Account_Created_On_Design_SS_Version' })
tableSample2= tableSample2.rename(index={'Account_Created_On_Design_SS_Version': ' ' })


plt.figure(figsize=(14, 8))
l = sns.countplot(x = "Account_Created_On_Design_SS_Version", hue = "Is_Product_Qualified_Lead?", data = InlineBA_True)
sns.despine()
plt.title("Is_Product_Qualified_Lead true version", fontsize = 18)
plt.xlabel("Account_Created_On_Design_SS_Version", fontsize = 14)
plt.ylabel("Count_User_ID", fontsize = 14)

for p in l.patches:
    l.annotate(p.get_height(), (p.get_x() + p.get_width() / 2., p.get_height() + 5), ha = 'center', va = 'center')

plt.savefig('PI/Product_Qualified_Lead true version.jpeg', dpi=100)

plt.figure(figsize=(14, 8))
l = sns.countplot(x = "Account_Created_On_Design_SS_Version", hue = "Is_Product_Qualified_Lead?", data = InlineBA_False)
sns.despine()
plt.title("Is_Product_Qualified_Lead true version", fontsize = 18)
plt.xlabel("Account_Created_On_Design_SS_Version", fontsize = 14)
plt.ylabel("Count_User_ID", fontsize = 14)

for p in l.patches:
    l.annotate(p.get_height(), (p.get_x() + p.get_width() / 2., p.get_height() + 5), ha = 'center', va = 'center')

plt.savefig('PI/Product_Qualified_Lead False version.jpeg', dpi=100)

tablesam1 = pd.pivot_table(InlineBA_True,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

tablesam2 = pd.pivot_table(InlineBA_False,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["User_ID"],
                       aggfunc={"User_ID": len},
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)


table_Timefalse = pd.pivot_table(InlineBA_False,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["secondsOnAssessmentStartPage", "secondsOnBaSelfConceptPage", "secondsOnBaSelfPage", "secondsOnCompletedPage", "totalSecondsSpentOnAssessment", ],
                       aggfunc={"secondsOnAssessmentStartPage": np.sum, "secondsOnBaSelfConceptPage": np.sum, "secondsOnBaSelfPage": np.sum, "secondsOnCompletedPage": np.sum, "totalSecondsSpentOnAssessment": np.sum},
                       fill_value=pd.Timedelta(hours=0),
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

table_TimeF = table_Timefalse.applymap(convert)


Table_TimeFN = table_TimeF.merge(tablesam2, on= 'Account_Created_On_Design_SS_Version', how= 'left')

table_Timetrue = pd.pivot_table(InlineBA_True,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["secondsOnAssessmentStartPage", "secondsOnBaSelfConceptPage", "secondsOnBaSelfPage", "secondsOnCompletedPage", "totalSecondsSpentOnAssessment", ],
                       aggfunc={"secondsOnAssessmentStartPage": np.sum, "secondsOnBaSelfConceptPage": np.sum, "secondsOnBaSelfPage": np.sum, "secondsOnCompletedPage": np.sum, "totalSecondsSpentOnAssessment": np.sum},
                       fill_value=pd.Timedelta(hours=0),
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

table_TimeT = table_Timetrue.applymap(convert)
Table_TimeTN = table_TimeT.merge(tablesam1, on= 'Account_Created_On_Design_SS_Version', how= 'left')

table_true = pd.pivot_table(InlineBA_True,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["secondsOnLoadingPage"],
                       aggfunc={"secondsOnLoadingPage": np.sum},
                       fill_value=pd.Timedelta(hours=0),
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

table_InlineT = table_true.applymap(convert)

table_false = pd.pivot_table(InlineBA_False,
                       index=['Account_Created_On_Design_SS_Version'],
                       #columns=["isInlineBaComplete"],
                       values=["secondsOnLoadingPage"],
                       aggfunc={"secondsOnLoadingPage": np.sum},
                       fill_value=pd.Timedelta(hours=0),
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)

table_InlineF = table_false.applymap(convert)



'''
Analysis as per Company perspective

'''

df['Users_per_Company']=df.groupby('CompanyName')['User_ID'].transform('count')


table_Company = pd.pivot_table(df,
                       index=['CompanyName'],
                       columns=["isInlineBaComplete"],
                       values=["Users_per_Company"],
                       aggfunc={"Users_per_Company": len},
                       dropna="False",
                       margins='True',
                       margins_name="x Grand Total")
                       #fill_value=0)



writer = pd.ExcelWriter('PI/PI_Analysis.xlsx')
df.to_excel(writer, sheet_name='Sheet1', index=False)
table.to_excel(writer, sheet_name= 'Num_Users_DesignV', index=True)
table2.to_excel(writer, sheet_name= 'Num_Users_complete_BA', index=True)
#InlineBA_True.to_excel(writer, sheet_name='Sheet4', index=False)
#InlineBA_False.to_excel(writer, sheet_name='Sheet5', index=False)
table3_.to_excel(writer, sheet_name= 'Time_spent_BA', index=True)
table4.to_excel(writer, sheet_name= 'User_True_BA', index=True)
True_table.to_excel(writer, sheet_name= 'User_True_BA%', index=True)
table5.to_excel(writer, sheet_name= 'User_False_BA', index=True)
False_table.to_excel(writer, sheet_name= 'User_False_BA%', index=True)
tableSample1.to_excel(writer, sheet_name= 'QualifiedInlineBA_False', index=True)
tableSample2.to_excel(writer, sheet_name= 'QualifiedInlineBA_True', index=True)
Table_TimeFN.to_excel(writer, sheet_name= 'TimeSpent_InlineBA_False', index=True)
Table_TimeTN.to_excel(writer, sheet_name= 'TimeSpent_InlineBA_True', index=True)
table_InlineT.to_excel(writer, sheet_name= 'Loading_InlineBA_True', index=True)
table_InlineF.to_excel(writer, sheet_name= 'Loading_InlineBA_False', index=True)
table_Company.to_excel(writer, sheet_name= 'UserBA_by_Company', index=True)
writer.save()

s.stop()
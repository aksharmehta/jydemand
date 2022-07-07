#--global.dataFrameSerialization="legacy"
from concurrent.futures import process
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid
from st_aggrid.grid_options_builder import GridOptionsBuilder
from st_aggrid.shared import GridUpdateMode
from st_aggrid.shared import JsCode
import pandas as pd
import numpy as np
from io import BytesIO
import sys
import time
import xlsxwriter



def aggrid_interactive_table(df: pd.DataFrame):
    """Creates an st-aggrid interactive table based on a dataframe.

    Args:
        df (pd.DataFrame]): Source dataframe

    Returns:
        dict: The selected row
    """
    options = GridOptionsBuilder.from_dataframe(
        df, enableRowGroup=True, enableValue=True, enablePivot=True
    )

    options.configure_side_bar()

    options.configure_selection("single")
    selection = AgGrid(
        df,
        enable_enterprise_modules=True,
        gridOptions=options.build(),
        theme="light",
        update_mode=GridUpdateMode.MODEL_CHANGED,
        allow_unsafe_jscode=True,
    )

    return selection


def _max_width_():
    max_width_str = f"max-width: 1800px;"
    st.markdown(
        f"""
    <style>
    .reportview-container .main .block-container{{
        {max_width_str}
    }}
    </style>    
    """,
        unsafe_allow_html=True,
    )


def pointerFiles(office, factory):
    #print("In pointerFiles")
    office.columns =[column.replace(" ", "_") for column in office.columns] 
    factory.columns =[column.replace(" ", "_") for column in factory.columns]
    dfP = pd.merge(office, factory, how = "left", left_on = "Flute_Bag", right_on = "Flute_Bag_No")
    emp = ['Priyank', 'Harshit', 'Atit', 'Darshil','Kunal']
    dfP = dfP.query('CTG_x=="D" and BAGGING_PRIORITIES.notnull() and Employee in @emp')
    filteredDF = dfP[['RmCode','Sz','Lt','Wdth','Total_Req(cts)','RmQty','Bag_No',	'Cust_Cd',	'Order_Type', 'Employee',	'BAGGING_PRIORITIES']].copy()
    #print("filtered DF is:")
    #print(filteredDF)
    return filteredDF


def processFiles(df1, df2, dfProv):
        df1.columns =[column.replace(" ", "_") for column in df1.columns] 
        df2.columns =[column.replace(" ", "_") for column in df2.columns] 
        dfProv.columns =[column.replace(" ", "_") for column in dfProv.columns] 
        dfProv = dfProv[["RmCode", "Sz","Lt","Wdth","Total_Req(cts)","Stock(cts)","Net(cts)","RmQty","BAGGING_PRIORITIES"]]
        dfProv.fillna(0, inplace=True)


        dfP = pd.merge(df1, df2, how = "left", left_on = "Flute_Bag", right_on = "Flute_Bag_No")

        #print(dfP.columns)
        dfP = dfP.query('CTG_x=="D" and BAGGING_PRIORITIES.notnull()')
        dfP.replace('ZSELF-ST', 'ZSELF', regex=True)
        dfP.to_pickle("dFp.pkl") 

        dfP.loc[dfP.Customer_Code == 'ZSELF', 'BAGGING_PRIORITIES'] = "ZSELF"
        dfP.loc[dfP.Customer_Code == 'ZSELF-ST', 'BAGGING_PRIORITIES'] = "ZSELF"

        #dfP.to_excel("Bagging List with Priority.xlsx")
        filteredDF1 = dfP[["RmCode", "Sz","Lt","Wdth","Total_Req(cts)","Stock(cts)","Net(cts)","RmQty","BAGGING_PRIORITIES"]].copy()
        filteredDF = filteredDF1.append(dfProv, ignore_index=True)
        print(filteredDF)

        filteredDF["BAGGING_PRIORITIES"] = filteredDF["BAGGING_PRIORITIES"].fillna("None")
        filteredDF["BAGGING_PRIORITIES"] = filteredDF["BAGGING_PRIORITIES"].astype(str)
        #filteredDF.loc[filteredDF["BAGGING_PRIORITIES"].isin(['1.0',1.0])]='1'
        #filteredDF.loc[filteredDF["BAGGING_PRIORITIES"].isin(['2.0',2.0])]='2'
        #filteredDF.loc[filteredDF["BAGGING_PRIORITIES"].isin(['3.0',3.0])]='3'
        #filteredDF.loc[filteredDF["BAGGING_PRIORITIES"].isin(['4.0',4.0])]='4'
        #filteredDF.loc[filteredDF["BAGGING_PRIORITIES"].isin(['5.0',5.0])]='5'
        listPriority = set(filteredDF["BAGGING_PRIORITIES"].tolist())
        
        listPriority1 = {'1+','ANAD','SJMG','ZSELF','1','2','3','4','5','6','PROVISION'}



        for i in listPriority1:
            add_df = {"RmCode":"IGNORE", "Sz":200,"Lt":1,"Wdth":1,"Total_Req(cts)":1,"Stock(cts)":5,"Net(cts)":2,"RmQty":10,"BAGGING_PRIORITIES": i}
            filteredDF = filteredDF.append(add_df,ignore_index = True)

        add_df = {"RmCode":"IGNORE", "Sz":200,"Lt":1,"Wdth":1,"Total_Req(cts)":1,"Stock(cts)":5,"Net(cts)":2,"RmQty":10,"BAGGING_PRIORITIES": '1+'}
        filteredDF = filteredDF.append(add_df,ignore_index = True)

        add_df2 = {"RmCode":"IGNORE", "Sz":200,"Lt":1,"Wdth":1,"Total_Req(cts)":1,"Stock(cts)":5,"Net(cts)":2,"RmQty":10,"BAGGING_PRIORITIES": '1+COD'}
        filteredDF = filteredDF.append(add_df2,ignore_index = True)

        add_df3 = {"RmCode":"IGNORE", "Sz":200,"Lt":1,"Wdth":1,"Total_Req(cts)":1,"Stock(cts)":5,"Net(cts)":2,"RmQty":10,"BAGGING_PRIORITIES": 'ZSELF'}
        filteredDF = filteredDF.append(add_df3,ignore_index = True)

        add_df4 = {"RmCode":"IGNORE", "Sz":200,"Lt":1,"Wdth":1,"Total_Req(cts)":1,"Stock(cts)":5,"Net(cts)":2,"RmQty":10,"BAGGING_PRIORITIES": 'ANAD'}
        filteredDF = filteredDF.append(add_df4,ignore_index = True)

        add_df5 = {"RmCode":"IGNORE", "Sz":200,"Lt":1,"Wdth":1,"Total_Req(cts)":1,"Stock(cts)":5,"Net(cts)":2,"RmQty":10,"BAGGING_PRIORITIES": 'SJMG'}
        filteredDF = filteredDF.append(add_df5,ignore_index = True)

        filteredDF["BAGGING_PRIORITIES"] = filteredDF["BAGGING_PRIORITIES"].fillna("None")

        listPriority = set(filteredDF["BAGGING_PRIORITIES"].tolist())

        stockP = set(filteredDF["Stock(cts)"].tolist())

        freq = filteredDF["Stock(cts)"].value_counts()
        items = freq[freq>1].index
        #print(items)




        #print(listPriority)

        #for value in listPriority:
        #   filteredDF[value] = 0

        #print(filteredDF["Stock(cts)"] / (filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]))
        filteredDF['p'] = filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]
        #filteredDF.loc[(filteredDF['Stock(cts)'] > 0) & (filteredDF['Stock(cts)'] in  items)) , 'StockPcs'] = filteredDF[(filteredDF['Stock(cts)'] > 0) & (filteredDF['Stock(cts)'] in  items)


        #if((filteredDF["p"]) < 0.01):
         #   filteredDF["StockPcs"] = filteredDF["Stock(cts)"] / ((filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]).round(3))
        #elif((filteredDF["p"]) < 0.1 & filteredDF['p'] > 0.09):
         #   filteredDF["StockPcs"] = filteredDF["Stock(cts)"] / ((filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]).round(2))
        #   filteredDF["StockPcs"] = filteredDF["Stock(cts)"] / ((filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]).round(1))

        filteredDF.loc[(filteredDF['p'] < 0.01) , 'StockPcs'] = filteredDF["Stock(cts)"] / ((filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]).round(3))
        filteredDF.loc[(filteredDF['p'] >= 0.01) & (filteredDF['p'] <= 0.09) , 'StockPcs'] = filteredDF["Stock(cts)"] / ((filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]).round(2))
        filteredDF.loc[(filteredDF['p'] >0.09) , 'StockPcs'] = filteredDF["Stock(cts)"] / ((filteredDF["Total_Req(cts)"]/filteredDF["RmQty"]).round(1))


        filteredDF["StockPcs"] = filteredDF["StockPcs"].round(0)
       
       

        #filteredDF("StockPcs") = filteredDF("StockPcs").asType(int).round(0)

        for value in listPriority:
            filteredDF.loc[filteredDF['BAGGING_PRIORITIES'] == value, str(value)] = filteredDF['RmQty']

        #print(listPriority)


        dfPivot = pd.pivot_table(filteredDF, index = ["RmCode", "Sz", "Lt", "Wdth", "StockPcs"], values = ['RmQty','1+COD','ANAD','SJMG','ZSELF','1+', '1','2','3','4','5','6','PROVISION'], aggfunc =np.sum)
        #print(dfPivot)
        new_order = ['RmQty', '1+COD','ANAD','SJMG','ZSELF','1+', '1','2','3','4','5','6','PROVISION']


        dfPivot = dfPivot.reindex(new_order,axis = 1)
        #dfPivot.style.apply(highlight_max, subset = ["1+","1"])

        dfPivot.columns = dfPivot.columns.tolist()
        #dfPivot["NetPcs"] = dfPivot["RmQty"] - dfPivot["StockPcs"]
        #print(dfPivot)

        unPivot = pd.DataFrame(dfPivot.to_records())
        #print(unPivot)

        unPivot["NetToBuy"] = unPivot["RmQty"] - unPivot["StockPcs"]
        unPivot["Diff"] = 0
        unPivot["Diff1"] = 0
        unPivot["Diff2"] = 0
        unPivot["Diff3"] = 0
        unPivot["Diff4"] = 0
        unPivot["Diff5"] = 0
        unPivot["Diff6"] = 0
        unPivot["Diff9"] = 0
        unPivot["Diff11"] = 0
        unPivot["Diff12"] = 0
        unPivot["Diff15"] = 0
        unPivot["DiffPROVISION"] = 0


        #print(unPivot.loc[unPivot['Sz'] == 0.04])


        unPivot.loc[unPivot['NetToBuy'] <= 0, ['RmQty', '1+COD','ANAD','SJMG','ZSELF','1+', "1","2","3","4","5","6","PROVISION"]] = 0

        unPivot.loc[unPivot['RmQty'] > 0, ['RmQty']] = unPivot['RmQty'] - unPivot['StockPcs']
        #unPivot.loc[unPivot['1+'] <=unPivot['RmQty'] , ['Diff','RmQty', '1+']] = [unPivot['RmQty'],unPivot["RmQty"]-unPivot['1+'],unPivot['1+']-unPivot["Diff"]]



        unPivot["Diff9"] = unPivot["StockPcs"]-unPivot["1+COD"]
        unPivot.loc[(unPivot['1+COD'] > 0) & (unPivot['StockPcs'] > 0), "1+COD"] = unPivot['1+COD'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['1+COD'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff9'] == unPivot['StockPcs']), "StockPcs"] = 0
        unPivot.loc[(unPivot['1+COD'] > 0) & (unPivot['Diff9'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff9']
        #New code
        unPivot.loc[(unPivot['1+COD'] < 0) & (unPivot['Diff9'] > 0),"StockPcs"] = unPivot['Diff9']
        unPivot.loc[(unPivot['1+COD'] < 0),"1+COD"] = 0
        unPivot.loc[(unPivot['1+COD'] > 0) & (unPivot['Diff9'] < 0),"StockPcs"] = 0


        unPivot["Diff11"] = unPivot["StockPcs"]-unPivot["ANAD"]
        unPivot.loc[(unPivot['ANAD'] > 0) & (unPivot['StockPcs'] > 0), "ANAD"] = unPivot['ANAD'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['ANAD'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff11'] == unPivot['StockPcs']), "StockPcs"] = 0
        unPivot.loc[(unPivot['ANAD'] > 0) & (unPivot['Diff11'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff11']
        #New code
        unPivot.loc[(unPivot['ANAD'] < 0) & (unPivot['Diff11'] > 0),"StockPcs"] = unPivot['Diff11']
        unPivot.loc[(unPivot['ANAD'] < 0),"ANAD"] = 0
        unPivot.loc[(unPivot['ANAD'] > 0) & (unPivot['Diff11'] < 0),"StockPcs"] = 0



        unPivot["Diff12"] = unPivot["StockPcs"]-unPivot["SJMG"]
        unPivot.loc[(unPivot['SJMG'] > 0) & (unPivot['StockPcs'] > 0), "SJMG"] = unPivot['SJMG'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['SJMG'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff12'] == unPivot['StockPcs']), "StockPcs"] = 0
        unPivot.loc[(unPivot['SJMG'] > 0) & (unPivot['Diff12'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff12']
        #New code
        unPivot.loc[(unPivot['SJMG'] < 0) & (unPivot['Diff12'] > 0),"StockPcs"] = unPivot['Diff12']
        unPivot.loc[(unPivot['SJMG'] < 0),"SJMG"] = 0
        unPivot.loc[(unPivot['SJMG'] > 0) & (unPivot['Diff12'] < 0),"StockPcs"] = 0

        unPivot["Diff15"] = unPivot["StockPcs"]-unPivot["ZSELF"]
        unPivot.loc[(unPivot['ZSELF'] > 0) & (unPivot['StockPcs'] > 0), "ZSELF"] = unPivot['ZSELF'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['ZSELF'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff15'] == unPivot['StockPcs']), "StockPcs"] = 0
        unPivot.loc[(unPivot['ZSELF'] > 0) & (unPivot['Diff15'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff15']
        #New code
        unPivot.loc[(unPivot['ZSELF'] < 0) & (unPivot['Diff15'] > 0),"StockPcs"] = unPivot['Diff15']
        unPivot.loc[(unPivot['ZSELF'] < 0),"ZSELF"] = 0
        unPivot.loc[(unPivot['ZSELF'] > 0) & (unPivot['Diff15'] < 0),"StockPcs"] = 0



        unPivot["Diff"] = unPivot["StockPcs"]-unPivot["1+"]
        unPivot.loc[(unPivot['1+'] > 0) & (unPivot['StockPcs'] > 0), "1+"] = unPivot['1+'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['1+'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff'] == unPivot['StockPcs']), "StockPcs"] = 0
        unPivot.loc[(unPivot['1+'] > 0) & (unPivot['Diff'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff']
        #New code
        unPivot.loc[(unPivot['1+'] < 0) & (unPivot['Diff'] > 0),"StockPcs"] = unPivot['Diff']
        unPivot.loc[(unPivot['1+'] < 0),"1+"] = 0
        unPivot.loc[(unPivot['1+'] > 0) & (unPivot['Diff'] < 0),"StockPcs"] = 0

        unPivot["Diff1"] = unPivot["StockPcs"]-unPivot["1"]
        unPivot.loc[(unPivot['1'] > 0) & (unPivot['StockPcs'] > 0), "1"] = unPivot['1'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['1'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff1'] == unPivot['StockPcs']), "StockPcs"] = 0
        #unPivot.loc[(unPivot['Diff1'] <= 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = 0
        unPivot.loc[(unPivot['1'] > 0) & (unPivot['Diff1'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff1']
        #New code
        unPivot.loc[(unPivot['1'] < 0) & (unPivot['Diff1'] > 0),"StockPcs"] = unPivot['Diff1']
        unPivot.loc[(unPivot['1'] < 0),"1"] = 0
        unPivot.loc[(unPivot['1'] > 0) & (unPivot['Diff1'] < 0),"StockPcs"] = 0

        unPivot["Diff2"] = unPivot["StockPcs"]-unPivot["2"]
        unPivot.loc[(unPivot['2'] > 0) & (unPivot['StockPcs'] > 0), "2"] = unPivot['2'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['2'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff2'] == unPivot['StockPcs']), "StockPcs"] = 0
        #unPivot.loc[(unPivot['Diff2'] <= 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = 0
        unPivot.loc[(unPivot['2'] > 0) & (unPivot['Diff2'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff2']
        #New code
        unPivot.loc[(unPivot['2'] < 0) & (unPivot['Diff2'] > 0),"StockPcs"] = unPivot['Diff2']
        unPivot.loc[(unPivot['2'] < 0),"2"] = 0
        unPivot.loc[(unPivot['2'] > 0) & (unPivot['Diff2'] < 0),"StockPcs"] = 0

        unPivot["Diff3"] = unPivot["StockPcs"]-unPivot["3"]
        unPivot.loc[(unPivot['3'] > 0) & (unPivot['StockPcs'] > 0), "3"] = unPivot['3'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['3'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff3'] == unPivot['StockPcs']), "StockPcs"] = 0
        #unPivot.loc[(unPivot['Diff3'] <= 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = 0
        unPivot.loc[(unPivot['3'] > 0) & (unPivot['Diff3'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff3']
        #New code
        unPivot.loc[(unPivot['3'] < 0) & (unPivot['Diff3'] > 0),"StockPcs"] = unPivot['Diff3']
        unPivot.loc[(unPivot['3'] < 0),"3"] = 0
        unPivot.loc[(unPivot['3'] > 0) & (unPivot['Diff3'] < 0),"StockPcs"] = 0


        unPivot["Diff4"] = unPivot["StockPcs"]-unPivot["4"]
        unPivot.loc[(unPivot['4'] > 0) & (unPivot['StockPcs'] > 0), "4"] = unPivot['4'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['4'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff4'] == unPivot['StockPcs']), "StockPcs"] = 0
        #unPivot.loc[(unPivot['Diff4'] <= 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = 0
        unPivot.loc[(unPivot['4'] > 0) & (unPivot['Diff4'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff4']
        #New code
        unPivot.loc[(unPivot['4'] < 0) & (unPivot['Diff4'] > 0),"StockPcs"] = unPivot['Diff4']
        unPivot.loc[(unPivot['4'] < 0),"4"] = 0
        unPivot.loc[(unPivot['4'] > 0) & (unPivot['Diff4'] < 0),"StockPcs"] = 0

        unPivot["Diff5"] = unPivot["StockPcs"]-unPivot["5"]
        unPivot.loc[(unPivot['5'] > 0) & (unPivot['StockPcs'] > 0), "5"] = unPivot['5'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['5'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff5'] == unPivot['StockPcs']), "StockPcs"] = 0
        #unPivot.loc[(unPivot['Diff5'] <= 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = 0
        unPivot.loc[(unPivot['5'] > 0) & (unPivot['Diff5'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff5']
        #New code
        unPivot.loc[(unPivot['5'] < 0) & (unPivot['Diff5'] > 0),"StockPcs"] = unPivot['Diff5']
        unPivot.loc[(unPivot['5'] < 0),"5"] = 0
        unPivot.loc[(unPivot['5'] > 0) & (unPivot['Diff5'] < 0),"StockPcs"] = 0


        unPivot["Diff6"] = unPivot["StockPcs"]-unPivot["6"]
        unPivot.loc[(unPivot['6'] > 0) & (unPivot['StockPcs'] > 0), "6"] = unPivot['6'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['6'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['Diff6'] == unPivot['StockPcs']), "StockPcs"] = 0
        #unPivot.loc[(unPivot['Diff6'] <= 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = 0
        unPivot.loc[(unPivot['6'] > 0) & (unPivot['Diff6'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['Diff6']
        #New code
        unPivot.loc[(unPivot['6'] < 0) & (unPivot['Diff6'] > 0),"StockPcs"] = unPivot['Diff6']
        unPivot.loc[(unPivot['6'] < 0),"6"] = 0
        unPivot.loc[(unPivot['6'] > 0) & (unPivot['Diff6'] < 0),"StockPcs"] = 0


        unPivot["DiffPROVISION"] = unPivot["StockPcs"]-unPivot["PROVISION"]
        unPivot.loc[(unPivot['PROVISION'] > 0) & (unPivot['StockPcs'] > 0), "PROVISION"] = unPivot['PROVISION'] - unPivot['StockPcs']
        unPivot.loc[(unPivot['PROVISION'] > 0) & unPivot['StockPcs'] > 0 & (unPivot['DiffPROVISION'] == unPivot['StockPcs']), "StockPcs"] = 0
        #unPivot.loc[(unPivot['DiffPROVISION'] <= 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = 0
        unPivot.loc[(unPivot['PROVISION'] > 0) & (unPivot['DiffPROVISION'] > 0) & (unPivot['StockPcs'] > 0), "StockPcs"] = unPivot['StockPcs'] - unPivot['DiffPROVISION']
        #New code
        unPivot.loc[(unPivot['PROVISION'] < 0) & (unPivot['DiffPROVISION'] > 0),"StockPcs"] = unPivot['DiffPROVISION']
        unPivot.loc[(unPivot['PROVISION'] < 0),"PROVISION"] = 0
        unPivot.loc[(unPivot['PROVISION'] > 0) & (unPivot['DiffPROVISION'] < 0),"StockPcs"] = 0



        unPivot1 = unPivot.query('RmQty > 0 and RmCode != "IGNORE"')
        pivoted = pd.pivot_table(unPivot1, index = ["RmCode", "Sz", "Lt", "Wdth", "StockPcs"], values = ['RmQty','1+COD','ANAD','SJMG','ZSELF','1+', '1','2','3','4','5','6','PROVISION'], aggfunc =np.sum)
        pivoted = pivoted.reindex(new_order,axis = 1)



        TodaysDate = time.strftime("%d-%m-%Y")
        excelfilename = "Demand" +TodaysDate +".xlsx"
        return pivoted


st.set_page_config(page_icon="✂️", page_title="Demand File Generator")

# st.image("https://emojipedia-us.s3.dualstack.us-west-1.amazonaws.com/thumbs/240/apple/285/balloon_1f388.png", width=100)
st.image(
    "http://www.kantilalchhotalal.com/wp-content/uploads/2021/03/logo2.png",
    width=100,
)

st.title("Demand File Generator")

# st.caption(
#     "PRD : TBC | Streamlit Ag-Grid from Pablo Fonseca: https://pypi.org/project/streamlit-aggrid/"
# )


# ModelType = st.radio(
#     "Choose your model",
#     ["Flair", "DistilBERT (Default)"],
#     help="At present, you can choose between 2 models (Flair or DistilBERT) to embed your text. More to come!",
# )

# with st.expander("ToDo's", expanded=False):
#     st.markdown(
#         """
# -   Add pandas.json_normalize() - https://streamlit.slack.com/archives/D02CQ5Z5GHG/p1633102204005500
# -   **Remove 200 MB limit and test with larger CSVs**. Currently, the content is embedded in base64 format, so we may end up with a large HTML file for the browser to render
# -   **Add an encoding selector** (to cater for a wider array of encoding types)
# -   **Expand accepted file types** (currently only .csv can be imported. Could expand to .xlsx, .txt & more)
# -   Add the ability to convert to pivot → filter → export wrangled output (Pablo is due to change AgGrid to allow export of pivoted/grouped data)
# 	    """
#     )
# 
#     st.text("")


c29, c30, c31 = st.columns([1, 6, 1])
#st.markdown(".stfile_uploader > label {font-size:105%; font-weight:bold; color:blue;} ",unsafe_allow_html=True)

with c30:

    uploaded_file = st.file_uploader(
        "Upload Office File",
        key="1",
        help="To activate 'wide mode', go to the hamburger menu > Settings > turn on 'wide mode'",
    )

    if uploaded_file is not None:
        file_container = st.expander("Check your uploaded OFFICE file")
        xl = pd.ExcelFile(uploaded_file)
        if len(xl.sheet_names) > 1:
            st.info("More than 1 sheet!")
            st.stop()
        office = pd.read_excel(uploaded_file)
        uploaded_file.seek(0)
        file_container.write(office)

    else:
        st.info(
            f"""
                👆 Upload the Office File here
                """
        )

        st.stop()


    uploaded_file2 = st.file_uploader(
        "Upload Factory File",
        key="2",
        help="To activate 'wide mode', go to the hamburger menu > Settings > turn on 'wide mode'",
    )

    if uploaded_file2 is not None:
        file_container = st.expander("Check your uploaded FACTORY file")
        xl = pd.ExcelFile(uploaded_file2)
        if len(xl.sheet_names) > 1:
            st.info("More than 1 sheet!")
            st.stop()
        factory = pd.read_excel(uploaded_file2, header = 1)
        uploaded_file2.seek(0)
        file_container.write(factory)

    else:
        st.info(
            f"""
                👆 Upload the Factory File here
                """
        )

        st.stop()
    
    uploaded_file3 = st.file_uploader(
        "Upload Provision File",
        key="3",
        help="To activate 'wide mode', go to the hamburger menu > Settings > turn on 'wide mode'",
    )

    if uploaded_file3 is not None:
        file_container = st.expander("Check your uploaded PROVISION file")
        xl = pd.ExcelFile(uploaded_file3)
        if len(xl.sheet_names) > 1:
            st.info("More than 1 sheet!")
            st.stop()
        provision = pd.read_excel(uploaded_file3)
        uploaded_file3.seek(0)
        file_container.write(provision)

    else:
        st.info(
            f"""
                👆 Upload the Provision File here
                """
        )

        st.stop()

from st_aggrid import GridUpdateMode, DataReturnMode

df = processFiles(office, factory, provision)
pointerDF = pointerFiles(office, factory)

output = BytesIO()
writer = pd.ExcelWriter(output, engine='xlsxwriter')

df.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Sheet_1")
pointerDF.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Pointers", index = False)
#workbook = writer.book
#worksheet = writer.sheets["Sheet_1","Pointers"]

writer.close()

output.seek(0)

st.download_button(
    label="Download Demand File",
    data=output.getvalue(),
    file_name="workbook.xlsx",
    mime="application/vnd.ms-excel")

df1 = df.reset_index()

selection = aggrid_interactive_table(df1)

if selection:
    st.write("You selected:")
    st.json(selection["selected_rows"])



gb = GridOptionsBuilder.from_dataframe(pointerDF)
# enables pivoting on all columns, however i'd need to change ag grid to allow export of pivoted/grouped data, however it select/filters groups
gb.configure_default_column(enablePivot=True, enableValue=True, enableRowGroup=True)
gb.configure_selection(selection_mode="multiple", use_checkbox=True)
gb.configure_side_bar()  # side_bar is clearly a typo :) should by sidebar
gridOptions = gb.build()

st.success(
    f"""
        💡 Tip! Hold the shift key when selecting rows to select multiple rows at once!
        """
)

response = AgGrid(
    pointerDF,
    gridOptions=gridOptions,
    enable_enterprise_modules=True,
    update_mode=GridUpdateMode.MODEL_CHANGED,
    data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
    fit_columns_on_grid_load=False,
)

sel = pd.DataFrame(response["selected_rows"])

st.subheader("Filtered data will appear below 👇 ")
st.text("")

st.table(sel)
st.text("")

c29, c30, c31 = st.columns([1, 1, 2])